"""
Virtualization Sizing Calculator
Ingests RVTools or Dell Live Optics exports and sizes a target VMware cluster.

Layout of this file:
  1. Constants
  2. Generic helpers
  3. Workbook loading (cached)
  4. Parsers  -> each source is mapped into ONE common "inventory" shape
  5. Sizing   -> pure math, no Streamlit (unit-testable)
  6. HTML report
  7. Streamlit UI
"""
import base64
import gzip
import hashlib
import io
import json
import math
import os
import sys
from datetime import datetime
from html import escape

import pandas as pd
import streamlit as st
import streamlit.components.v1 as components

# =============================================================================
# 1. CONSTANTS
# =============================================================================
APP_TITLE = "Virtualization Sizing Calculator"
APP_VERSION = "3.3"
APP_DIR = os.path.dirname(os.path.abspath(__file__))
DEFAULT_LOGO_PATH = os.path.join(APP_DIR, "ahead_logo.png")          # navy, light backgrounds
WHITE_LOGO_PATH = os.path.join(APP_DIR, "ahead_logo_white.png")      # reversed, dark backgrounds

ALL_CLUSTERS = "All Clusters"
LIC_MIN_CORES_PER_SOCKET = 16   # VMware per-core licensing minimum per CPU
PERF_TARGET_UTIL = 0.80         # Performance sizing keeps hosts at <= 80% CPU
LO_95TH_FALLBACK = 0.95         # 95th estimate = 95% of peak when column missing
MIB_PER_TB = 1024 * 1024
EPS = 1e-9                      # guards ceil() against float noise (e.g. 3.0000000001)

# Hypervisor vCPU:pCPU guidance: (ratio, "average"/"maximum", source note)
HYPERVISORS = {
    "VMware": (4.5, "average", "Average for the workloads Intel has captured."),
    "Nutanix": (3.5, "average", "Average vCPU:pCPU ratio."),
    "Azure Local": (2.5, "maximum", "Maximum Intel recommends; Dell product management agrees."),
    "Hyper-V": (2.0, "average", "Average vCPU:pCPU ratio."),
}
DEFAULT_HYPERVISOR = "VMware"

RVTOOLS_SHEETS = ["vInfo", "vHost", "vDatastore", "vPartition", "vDisk", "vLicense"]
LIVEOPTICS_SHEETS = ["VMs", "ESX Hosts", "Host Devices", "ESX Performance", "ESX Licenses"]

HOST_NAME_COLS = ["Host Name", "Host", "ESX Host", "Hostname", "Server Name"]
LUN_ID_COLS = ["Canonical Name", "Device ID", "Device UUID", "UUID", "NAA ID", "Device Name"]


# =============================================================================
# 2. GENERIC HELPERS
# =============================================================================
def num(series):
    return pd.to_numeric(series, errors="coerce")


def num_col(df, col):
    """Numeric column, or an all-NaN series if the column is missing."""
    if col in df.columns:
        return num(df[col])
    return pd.Series(float("nan"), index=df.index, dtype="float64")


def col_sum(df, col):
    if df is None or col not in df.columns:
        return 0.0
    total = num(df[col]).sum()
    return float(total) if pd.notna(total) else 0.0


def first_col(df, candidates):
    for c in candidates:
        if c in df.columns:
            return c
    return None


def rvtools_tb(df, base):
    """RVTools: find '{base} <unit>' and normalise to TB."""
    for unit, divisor in (("MiB", MIB_PER_TB), ("MB", MIB_PER_TB),
                          ("GiB", 1024), ("GB", 1024), ("TiB", 1), ("TB", 1)):
        col = f"{base} {unit}"
        if col in df.columns:
            return col_sum(df, col) / divisor
    return 0.0


def is_true(series):
    return series.astype(str).str.strip().str.lower().isin(["true", "1", "yes"])


def mode_or(series, default=0):
    s = series.dropna()
    return s.mode().iloc[0] if not s.empty else default


def filter_clusters(df, clusters, col="Cluster"):
    """clusters=None means 'all'. Always returns a copy (safe to add columns)."""
    if df is None:
        return pd.DataFrame()
    if clusters is None or col not in df.columns:
        return df.copy()
    return df[df[col].astype(str).isin(clusters)].copy()


def powered_on_only(df, col):
    if col not in df.columns:
        return df
    return df[df[col].astype(str).str.contains("poweredOn", case=False, na=False)]


def power_counts(df, col):
    """(powered_on, powered_off) for VMs in scope, or (None, None) if no power column."""
    if col not in df.columns:
        return None, None
    on = int(df[col].astype(str).str.contains("poweredOn", case=False, na=False).sum())
    return on, len(df) - on


def cpu_ghz(df, mhz_cols, ghz_cols):
    c = first_col(df, mhz_cols)
    if c:
        return num(df[c]) / 1000
    c = first_col(df, ghz_cols)
    return num(df[c]) if c else pd.Series(float("nan"), index=df.index, dtype="float64")


def text_col(df, candidates):
    c = first_col(df, candidates)
    return df[c].astype(str) if c else pd.Series("Unknown", index=df.index)


def drop_templates(df):
    for col in ("Template", "SRM Placeholder"):
        if col in df.columns:
            df = df[~is_true(df[col])]
    return df


def ceil_div(a, b):
    if a <= 0 or b <= 0:
        return 0
    return math.ceil(a / b - EPS)


def fmt_ratio(x):
    return f"{x:.1f}:1" if x is not None else "n/a"


def empty_storage(note="Not available"):
    return {"cap": 0.0, "used": 0.0, "free": 0.0, "note": note}


# =============================================================================
# 3. WORKBOOK LOADING (cached - only re-runs when a new file is uploaded)
# =============================================================================
@st.cache_data(show_spinner="Reading workbook...", max_entries=10)
def load_workbook(file_bytes):
    """Returns (source_type, {sheet_name: DataFrame}). Only reads needed sheets."""
    xls = pd.ExcelFile(io.BytesIO(file_bytes), engine="openpyxl")
    names = set(xls.sheet_names)
    if "vInfo" in names:
        source, wanted = "RVTools", RVTOOLS_SHEETS
    elif {"VMs", "ESX Hosts"} <= names:
        source, wanted = "LiveOptics", LIVEOPTICS_SHEETS
    else:
        return None, {}

    sheets = {}
    for name in wanted:
        if name in names:
            df = xls.parse(name)
            df.columns = df.columns.astype(str).str.strip()
            sheets[name] = df
    return source, sheets


def cluster_list(source, sheets):
    df = sheets["vInfo"] if source == "RVTools" else sheets["VMs"]
    if "Cluster" not in df.columns:
        return []
    return sorted(df["Cluster"].dropna().astype(str).unique())


# =============================================================================
# 3b. PROJECT LIBRARY (saved projects keep ONLY the columns the sizing needs)
# =============================================================================
PROJECT_FORMAT = "hostsizer-project"
PROJECT_VERSION = 1
PROJECT_EXT = ".hsproj"
PROJECTS_DIR = os.path.expanduser(os.environ.get("HOSTSIZER_PROJECTS", os.path.join("~", "HostSizer", "Projects")))


def _unit_cols(*bases):
    units = ("MiB", "MB", "GiB", "GB", "TiB", "TB")
    return [f"{b} {u}" for b in bases for u in units]


# Every column the parsers can read. Anything else in the export is dropped when a project is saved.
# Keep this in sync with parse_rvtools() / parse_live_optics(); save_project() verifies nothing is lost.
REQUIRED_COLUMNS = {
    "RVTools": {
        "vInfo": ["VM", "Powerstate", "Template", "SRM Placeholder", "Cluster", "CPUs", "Memory", "Memory GB"]
                 + _unit_cols("Provisioned", "In Use"),
        "vHost": ["Cluster", "# CPU", "Cores per CPU", "# Memory", "CPU Model", "CPU model",
                  "Speed", "CPU Speed", "Speed MHz", "ESX Version", "Product"],
        "vDatastore": ["Name", "Cluster name", "Cluster"] + _unit_cols("Capacity", "In Use"),
        "vPartition": ["VM"] + _unit_cols("Consumed"),
        "vDisk": ["VM", "Raw"] + _unit_cols("Capacity"),
        "vLicense": ["Name", "Used"],
    },
    "LiveOptics": {
        "VMs": ["VM Name", "Power State", "Template", "SRM Placeholder", "Cluster", "Virtual CPU",
                "Provisioned Memory (MiB)", "Virtual Disk Size (MiB)", "Virtual Disk Used (MiB)",
                "Guest VM Disk Used (MiB)"],
        "ESX Hosts": ["Cluster", "CPU Sockets", "CPU Cores", "Memory (KiB)", "CPU Model", "Processor", "CPU Type",
                      "CPU Description", "CPU Speed (MHz)", "Speed (MHz)", "CPU Clock (MHz)", "CPU Speed (GHz)",
                      "Speed (GHz)", "ESX Version", "OS Version", "Version", "OS"] + HOST_NAME_COLS,
        "Host Devices": ["Cluster", "Capacity (GiB)", "Used Capacity (GiB)", "Free Capacity (GiB)"]
                        + HOST_NAME_COLS + LUN_ID_COLS,
        "ESX Performance": ["Cluster", "95th Percentile CPU (GHz)", "Peak CPU (GHz)", "Average CPU (GHz)"],
        "ESX Licenses": ["Software Title"],
    },
}
VERIFY_KEYS = ["tot_vms", "tot_vcpu", "tot_ram", "cur_host_count", "cur_cores", "cur_total_ram_gb", "hosts_now",
               "hosts_fut", "cur_lic_cores", "vm_prov_tb", "vm_used_tb", "guest_used_tb", "rdm_tb", "perf_ghz"]


def reduce_sheets(source, sheets):
    """Keep only the tabs and columns the sizing reads."""
    out = {}
    for name, cols in REQUIRED_COLUMNS[source].items():
        if name in sheets:
            df = sheets[name]
            out[name] = df[[c for c in dict.fromkeys(cols) if c in df.columns]].reset_index(drop=True)
    return out


def project_slug(name):
    slug = "".join(ch if ch.isalnum() or ch in " -_." else "_" for ch in name).strip(" .")
    return (slug or "project")[:120]


def project_path(name, folder=None):
    return os.path.join(folder or PROJECTS_DIR, project_slug(name) + PROJECT_EXT)


def frames_to_json(sheets):
    return {name: json.loads(df.to_json(orient="split", index=False, date_format="iso")) for name, df in sheets.items()}


def frames_from_json(data):
    frames = {}
    for name, split in data.items():
        df = pd.DataFrame(split["data"], columns=split["columns"])
        df.columns = df.columns.astype(str)
        frames[name] = df
    return frames


def write_project(path, payload):
    """Atomic write of a gzip-compressed JSON project file."""
    os.makedirs(os.path.dirname(path), exist_ok=True)
    tmp = path + ".tmp"
    with gzip.open(tmp, "wt", encoding="utf-8") as f:
        json.dump(payload, f)
    os.replace(tmp, path)


def read_project(path):
    with gzip.open(path, "rt", encoding="utf-8") as f:
        payload = json.load(f)
    if payload.get("format") != PROJECT_FORMAT:
        raise ValueError("Not a Host Sizer project file.")
    if payload.get("version", 0) > PROJECT_VERSION:
        raise ValueError("This project was saved by a newer version of Host Sizer.")
    return payload


@st.cache_data(show_spinner=False, max_entries=200)
def project_meta(path, mtime):
    """Light metadata for the Home tiles (cached until the file changes)."""
    p = read_project(path)
    return {k: p.get(k) for k in ("name", "saved", "created", "app_version", "summary")} | {
        "source_type": p["source"]["type"], "file_name": p["source"]["file_name"]}


def list_projects(folder=None):
    folder = folder or PROJECTS_DIR
    if not os.path.isdir(folder):
        return []
    items = []
    for fn in os.listdir(folder):
        if not fn.endswith(PROJECT_EXT):
            continue
        path = os.path.join(folder, fn)
        try:
            meta = project_meta(path, os.path.getmtime(path))
        except Exception as exc:  # noqa: BLE001 - show unreadable files instead of hiding them
            meta = {"name": fn, "error": str(exc)}
        items.append((path, meta))
    return sorted(items, key=lambda x: x[1].get("saved") or "", reverse=True)


# =============================================================================
# 4. PARSERS
# Each parser returns the same "inventory" dict:
#   vms   : DataFrame[name, vcpu, ram_gb]          (in scope)
#   hosts : DataFrame[sockets, cores, ram_gb]      (cores = total per host)
#   storage, vm_prov_tb, vm_used_tb, guest_used_tb, rdm_cnt, rdm_tb
#   has_perf, perf_ghz, perf_estimated, lic_edition, esxi_version
#   raw_vms, raw_hosts (original rows in scope, for the Raw Data tab)
# =============================================================================
def parse_rvtools(sheets, clusters, include_off):
    info = drop_templates(sheets["vInfo"])
    info = filter_clusters(info, clusters)
    vm_on, vm_off = power_counts(info, "Powerstate")
    if not include_off:
        info = powered_on_only(info, "Powerstate")

    if "Memory" in info.columns:            # RVTools 'Memory' is MiB
        ram_gb = num(info["Memory"]) / 1024
    else:
        ram_gb = num_col(info, "Memory GB")

    vms = pd.DataFrame({
        "name": info["VM"].astype(str) if "VM" in info.columns else "N/A",
        "vcpu": num_col(info, "CPUs"),
        "ram_gb": ram_gb,
    }, index=info.index)

    # Hosts
    vhost = filter_clusters(sheets.get("vHost"), clusters)
    sockets = num_col(vhost, "# CPU")
    hosts = pd.DataFrame({
        "sockets": sockets,
        "cores": sockets * num_col(vhost, "Cores per CPU"),
        "ram_gb": num_col(vhost, "# Memory") / 1024,
        "model": text_col(vhost, ["CPU Model", "CPU model"]),
        "ghz": cpu_ghz(vhost, ["Speed", "CPU Speed", "Speed MHz"], []),
    }, index=vhost.index)

    ver_col = first_col(vhost, ["ESX Version", "Product"])
    esxi_version = str(mode_or(vhost[ver_col], "Unknown")) if ver_col else "Unknown"

    # License edition (vLicense is vCenter-wide, not per cluster)
    lic_edition = "Unknown"
    lic = sheets.get("vLicense")
    if lic is not None and "Name" in lic.columns:
        lic = lic[~lic["Name"].astype(str).str.contains("vCenter", case=False, na=False)]
        if not lic.empty:
            used = num_col(lic, "Used").fillna(0)
            lic_edition = str(lic.loc[used.idxmax(), "Name"])

    # Datastores (shared, non-local)
    storage = empty_storage("Datastore data not found")
    dsdf = sheets.get("vDatastore")
    if dsdf is not None and not dsdf.empty:
        d = dsdf
        if "Name" in d.columns:
            d = d[~d["Name"].astype(str).str.contains("local", case=False, na=False)]
        note = "All shared datastores"
        if clusters is not None:
            ccol = first_col(d, ["Cluster name", "Cluster"])
            if ccol:
                wanted = set(clusters)
                d = d[d[ccol].astype(str).apply(
                    lambda s: any(c.strip() in wanted for c in s.split(",")))]
                note = "Filtered to selected clusters"
            else:
                note = "Global - datastores not mapped to clusters"
        cap, used = rvtools_tb(d, "Capacity"), rvtools_tb(d, "In Use")
        storage = {"cap": cap, "used": used, "free": cap - used, "note": note}

    vm_names = set(vms["name"])

    guest_used_tb = 0.0
    part = sheets.get("vPartition")
    if part is not None and "VM" in part.columns:
        guest_used_tb = rvtools_tb(part[part["VM"].astype(str).isin(vm_names)], "Consumed")

    rdm_cnt, rdm_tb = 0, 0.0
    disk = sheets.get("vDisk")
    if disk is not None and "Raw" in disk.columns:
        if "VM" in disk.columns:
            disk = disk[disk["VM"].astype(str).isin(vm_names)]
        rdm = disk[is_true(disk["Raw"])]
        rdm_cnt, rdm_tb = len(rdm), rvtools_tb(rdm, "Capacity")

    return {
        "vms": vms, "hosts": hosts, "storage": storage,
        "vm_prov_tb": rvtools_tb(info, "Provisioned"),
        "vm_used_tb": rvtools_tb(info, "In Use"),
        "guest_used_tb": guest_used_tb, "rdm_cnt": rdm_cnt, "rdm_tb": rdm_tb,
        "has_perf": False, "perf_ghz": 0.0, "perf_estimated": False,
        "lic_edition": lic_edition, "esxi_version": esxi_version,
        "vm_on": vm_on, "vm_off": vm_off, "include_off": include_off,
        "raw_vms": info, "raw_hosts": vhost,
    }


def parse_live_optics(sheets, clusters, include_off, basis):
    vm = drop_templates(sheets["VMs"])
    vm = filter_clusters(vm, clusters)
    vm_on, vm_off = power_counts(vm, "Power State")
    if not include_off:
        vm = powered_on_only(vm, "Power State")

    vms = pd.DataFrame({
        "name": vm["VM Name"].astype(str) if "VM Name" in vm.columns else "N/A",
        "vcpu": num_col(vm, "Virtual CPU"),
        "ram_gb": num_col(vm, "Provisioned Memory (MiB)") / 1024,
    }, index=vm.index)

    esx = filter_clusters(sheets.get("ESX Hosts"), clusters)
    hosts = pd.DataFrame({
        "sockets": num_col(esx, "CPU Sockets"),
        "cores": num_col(esx, "CPU Cores"),
        "ram_gb": num_col(esx, "Memory (KiB)") / 1024 / 1024,
        "model": text_col(esx, ["CPU Model", "Processor", "CPU Type", "CPU Description"]),
        "ghz": cpu_ghz(esx, ["CPU Speed (MHz)", "Speed (MHz)", "CPU Clock (MHz)"], ["CPU Speed (GHz)", "Speed (GHz)"]),
    }, index=esx.index)

    ver_col = first_col(esx, ["ESX Version", "OS Version", "Version", "OS"])
    esxi_version = str(mode_or(esx[ver_col], "Unknown")) if ver_col else "Unknown"

    # Storage from Host Devices: scope to hosts in cluster, de-duplicate shared LUNs
    storage = empty_storage("Host Devices data not found")
    dev = sheets.get("Host Devices")
    if dev is not None and not dev.empty:
        d, notes = dev, []
        if clusters is not None:
            if "Cluster" in d.columns:
                d = filter_clusters(d, clusters)
            else:
                hc_dev, hc_esx = first_col(d, HOST_NAME_COLS), first_col(esx, HOST_NAME_COLS)
                if hc_dev and hc_esx:
                    d = d[d[hc_dev].astype(str).isin(esx[hc_esx].astype(str))]
                else:
                    notes.append("not filtered by cluster")
        id_col = first_col(d, LUN_ID_COLS)
        if id_col:
            d = d.drop_duplicates(subset=[id_col])
            notes.insert(0, f"LUNs de-duplicated by '{id_col}'")
        else:
            notes.insert(0, "shared LUNs may be counted once per host")
        storage = {
            "cap": col_sum(d, "Capacity (GiB)") / 1024,
            "used": col_sum(d, "Used Capacity (GiB)") / 1024,
            "free": col_sum(d, "Free Capacity (GiB)") / 1024,
            "note": "Host Devices - " + "; ".join(notes),
        }

    # Performance
    has_perf, perf_ghz, perf_estimated = False, 0.0, False
    perf = sheets.get("ESX Performance")
    if perf is not None:
        p = filter_clusters(perf, clusters)
        if basis == "95th Percentile":
            if "95th Percentile CPU (GHz)" in p.columns:
                perf_ghz = col_sum(p, "95th Percentile CPU (GHz)")
            else:
                perf_ghz = col_sum(p, "Peak CPU (GHz)") * LO_95TH_FALLBACK
                perf_estimated = True
        elif basis == "Peak CPU":
            perf_ghz = col_sum(p, "Peak CPU (GHz)")
        else:
            perf_ghz = col_sum(p, "Average CPU (GHz)")
        has_perf = True

    lic_edition = "Unknown"
    lic = sheets.get("ESX Licenses")
    if lic is not None and "Software Title" in lic.columns:
        lic_edition = str(mode_or(lic["Software Title"], "Unknown"))

    return {
        "vms": vms, "hosts": hosts, "storage": storage,
        "vm_prov_tb": col_sum(vm, "Virtual Disk Size (MiB)") / MIB_PER_TB,
        "vm_used_tb": col_sum(vm, "Virtual Disk Used (MiB)") / MIB_PER_TB,
        "guest_used_tb": col_sum(vm, "Guest VM Disk Used (MiB)") / MIB_PER_TB,
        "rdm_cnt": 0, "rdm_tb": 0.0,
        "has_perf": has_perf, "perf_ghz": perf_ghz, "perf_estimated": perf_estimated,
        "lic_edition": lic_edition, "esxi_version": esxi_version,
        "vm_on": vm_on, "vm_off": vm_off, "include_off": include_off,
        "raw_vms": vm, "raw_hosts": esx,
    }


# =============================================================================
# 5. SIZING (pure functions - no Streamlit)
# =============================================================================
def summarize_hosts(hosts):
    out = {"cur_host_count": len(hosts), "cur_cores": 0.0, "cur_total_ram_gb": 0.0,
           "cur_numa_cores": 0.0, "cur_numa_ram": 0.0, "cur_lic_cores": 0,
           "cur_sockets": 0, "cur_ram_per_host": 0.0, "cur_max_host_cores": 0.0,
           "cur_cpu_model": "Unknown", "cur_ghz": 0.0, "cur_ghz_total": 0.0}
    if hosts.empty:
        return out
    sockets = hosts["sockets"].where(hosts["sockets"] > 0)
    cores_per_socket = hosts["cores"] / sockets
    out["cur_cores"] = float(hosts["cores"].sum())
    out["cur_total_ram_gb"] = float(hosts["ram_gb"].sum())
    out["cur_numa_cores"] = float(mode_or(cores_per_socket))
    out["cur_numa_ram"] = float(mode_or((hosts["ram_gb"] / sockets).round(0)))
    out["cur_lic_cores"] = int((sockets * cores_per_socket.clip(lower=LIC_MIN_CORES_PER_SOCKET)).sum())
    out["cur_sockets"] = int(mode_or(sockets))
    out["cur_ram_per_host"] = float(mode_or(hosts["ram_gb"].round(0)))
    out["cur_max_host_cores"] = float(hosts["cores"].max()) if hosts["cores"].notna().any() else 0.0
    models = hosts["model"][hosts["model"].str.lower().ne("unknown") & hosts["model"].str.lower().ne("nan")]
    out["cur_cpu_model"] = " ".join(str(mode_or(models, "Unknown")).split())
    if hosts["ghz"].notna().any():
        out["cur_ghz"] = float(mode_or(hosts["ghz"].round(2)))
        out["cur_ghz_total"] = float((hosts["cores"] * hosts["ghz"]).sum())
    return out


def summarize_vms(vms, numa_cores, numa_ram):
    out = {"tot_vms": len(vms), "tot_vcpu": float(vms["vcpu"].sum()), "tot_ram": float(vms["ram_gb"].sum()),
           "max_vm_cpu": 0, "name_max_cpu": "N/A", "max_vm_ram": 0.0, "name_max_ram": "N/A",
           "wide_cpu_count": int((vms["vcpu"] > numa_cores).sum()),
           "wide_ram_count": int((vms["ram_gb"] > numa_ram).sum())}
    if vms["vcpu"].notna().any():
        i = vms["vcpu"].idxmax()
        out["max_vm_cpu"], out["name_max_cpu"] = int(vms.at[i, "vcpu"]), str(vms.at[i, "name"])
    if vms["ram_gb"].notna().any():
        i = vms["ram_gb"].idxmax()
        out["max_vm_ram"], out["name_max_ram"] = float(vms.at[i, "ram_gb"]), str(vms.at[i, "name"])
    return out


def license_cores_per_host(sockets, cores_per_socket):
    return sockets * max(cores_per_socket, LIC_MIN_CORES_PER_SOCKET)


def design_status(r, hosts, vcpu, ram):
    """Failover load for a host count against the design rules."""
    active = hosts - r["ha_nodes"]
    usable_ram = max(active, 0) * r["eff_ram"]
    cpu_ratio = vcpu / (active * r["eff_cores"]) if active > 0 and r["eff_cores"] else 0.0
    ram_use = ram / usable_ram if usable_ram else 0.0
    return {"active": active, "usable_ram": usable_ram, "cpu_ratio": cpu_ratio, "ram_use": ram_use,
            "cpu_ok": active > 0 and cpu_ratio <= r["vcpu_ratio"] + EPS,
            "ram_ok": active > 0 and ram_use <= 1 + EPS,
            "ha_ok": active > 0, "min_ok": hosts >= r["min_hosts"]}


def over_hv(r, ratio):
    """True when a vCPU:pCPU ratio is above the selected hypervisor's guidance."""
    return ratio is not None and ratio > r["hv_ratio"] + EPS


def hv_label(r):
    return f"{r['hypervisor']} {r['hv_kind']} {r['hv_ratio']:.1f}:1"


def capacity_checks(r, hosts, vcpu, ram, calc_hosts):
    """Validate a (possibly forced) host count against the design rules.
    Returns list of (is_problem, text)."""
    ha, out = r["ha_nodes"], []
    active = hosts - ha
    hl = f"{hosts} host" + ("s" if hosts != 1 else "")
    if active <= 0:
        return [(True, f"{hl} leaves no active capacity with N+{ha} HA.")]
    if hosts < r["min_hosts"]:
        out.append((True, f"{hl} is below the {r['min_hosts']}-host minimum cluster size."))
    cpu_ratio = vcpu / (active * r["eff_cores"]) if r["eff_cores"] else 0
    if cpu_ratio > r["vcpu_ratio"] + EPS:
        out.append((True, f"CPU: {cpu_ratio:.1f}:1 vCPU per usable core during HA failover exceeds the "
                          f"{r['vcpu_ratio']:.1f}:1 design limit."))
    ram_use = ram / (active * r["eff_ram"]) if r["eff_ram"] else 0
    if ram_use > 1 + EPS:
        out.append((True, f"RAM: workload needs {ram_use:.0%} of usable RAM during HA failover "
                          f"({ram:,.0f} GB vs {active * r['eff_ram']:,.0f} GB)."))
    if not out:
        extra = hosts - calc_hosts
        tail = f" ({extra} host{'s' if extra != 1 else ''} above the calculated {calc_hosts})." if extra > 0 else "."
        out.append((False, f"{hl} meets the CPU and RAM design limits with N+{ha} HA{tail} "
                           f"Failover load: {cpu_ratio:.1f}:1 vCPU per usable core, {ram_use:.0%} of usable RAM."))
    return out


def compute_sizing(inv, p):
    """inv: inventory dict from a parser. p: parameter dict from the sidebar.
    Returns one flat results dict used by BOTH the on-screen view and the HTML report."""
    host_cores = p["sockets"] * p["cores"]
    eff_cores = host_cores * (1 - p["cpu_buffer"] / 100)
    eff_ram = p["ram"] * (1 - p["ram_buffer"] / 100)
    numa_cores, numa_ram = p["cores"], p["ram"] / p["sockets"]
    ha, min_hosts = p["ha_nodes"], p["min_hosts"]

    r = dict(p)
    r.update(summarize_hosts(inv["hosts"]))
    r.update(summarize_vms(inv["vms"], numa_cores, numa_ram))
    r.update({k: inv[k] for k in ("vm_prov_tb", "vm_used_tb", "guest_used_tb", "rdm_cnt", "rdm_tb",
                                  "has_perf", "perf_ghz", "perf_estimated", "lic_edition", "esxi_version",
                                  "vm_on", "vm_off")})
    r["storage"] = inv["storage"]
    hv = p.get("hypervisor") or DEFAULT_HYPERVISOR
    r["hypervisor"] = hv
    r["hv_ratio"], r["hv_kind"], r["hv_note"] = HYPERVISORS.get(hv, HYPERVISORS[DEFAULT_HYPERVISOR])
    r.update(host_cores=host_cores, eff_cores=eff_cores, eff_ram=eff_ram,
             tgt_numa_cores=numa_cores, tgt_numa_ram=numa_ram)

    def size(vcpu, ram):
        h_cpu = ceil_div(vcpu / p["vcpu_ratio"], eff_cores)
        h_ram = ceil_div(ram, eff_ram)
        constraint = "CPU" if h_cpu > h_ram else ("RAM" if h_ram > h_cpu else "Balanced")
        raw = max(h_cpu, h_ram)
        total = max(raw + ha, min_hosts)
        return h_cpu, h_ram, raw, constraint, total, (raw + ha) < min_hosts

    def ratios(vcpu, hosts):
        all_hosts = vcpu / (hosts * host_cores) if hosts > 0 else 0.0
        failover = vcpu / ((hosts - ha) * host_cores) if hosts - ha > 0 else None
        return all_hosts, failover

    # Current
    r["hosts_cpu"], r["hosts_ram"], r["raw_hosts"], r["constraint"], r["hosts_now"], r["min_applied_now"] = \
        size(r["tot_vcpu"], r["tot_ram"])
    r["calc_hosts_now"] = r["hosts_now"]
    r["override_now"] = bool(p.get("override"))
    if r["override_now"]:
        r["hosts_now"] = int(p["override_hosts"])
    r["ratio_now_all"], r["ratio_now_failover"] = ratios(r["tot_vcpu"], r["hosts_now"])
    r["cur_ratio"] = r["tot_vcpu"] / r["cur_cores"] if r["cur_cores"] > 0 else 0.0
    n1_cores = r["cur_cores"] - r["cur_max_host_cores"]
    r["cur_ratio_n1"] = r["tot_vcpu"] / n1_cores if r["cur_host_count"] > 1 and n1_cores > 0 else None
    r["cur_mem_alloc"] = r["tot_ram"] / r["cur_total_ram_gb"] if r["cur_total_ram_gb"] > 0 else None
    n1_ram = r["cur_total_ram_gb"] - r["cur_ram_per_host"]
    r["cur_mem_alloc_n1"] = r["tot_ram"] / n1_ram if r["cur_host_count"] > 1 and n1_ram > 0 else None
    r["cur_vms_per_host"] = r["tot_vms"] / r["cur_host_count"] if r["cur_host_count"] else None
    r["avg_vcpu"] = r["tot_vcpu"] / r["tot_vms"] if r["tot_vms"] else 0.0
    r["avg_ram"] = r["tot_ram"] / r["tot_vms"] if r["tot_vms"] else 0.0
    r["cur_cpu_util"] = (r["perf_ghz"] / r["cur_ghz_total"]
                         if r["has_perf"] and r["cur_ghz_total"] > 0 else None)

    # Future (compound growth)
    mult = (1 + p["growth"]) ** p["years"]
    r["fut_vcpu"], r["fut_ram"] = r["tot_vcpu"] * mult, r["tot_ram"] * mult
    r["fut_hosts_cpu"], r["fut_hosts_ram"], r["fut_raw_hosts"], r["fut_constraint"], r["hosts_fut"], r["min_applied_fut"] = \
        size(r["fut_vcpu"], r["fut_ram"])
    r["calc_hosts_fut"] = r["hosts_fut"]
    r["override_fut"] = r["override_now"] and bool(p.get("override_future", True))
    if r["override_fut"]:
        r["hosts_fut"] = int(p["override_hosts"])
    r["ratio_fut_all"], r["ratio_fut_failover"] = ratios(r["fut_vcpu"], r["hosts_fut"])

    # Design-rule validation (most useful when a host count is forced)
    r["checks_now"] = capacity_checks(r, r["hosts_now"], r["tot_vcpu"], r["tot_ram"], r["calc_hosts_now"])
    r["checks_fut"] = capacity_checks(r, r["hosts_fut"], r["fut_vcpu"], r["fut_ram"], r["calc_hosts_fut"])
    r["design_now"] = design_status(r, r["hosts_now"], r["tot_vcpu"], r["tot_ram"])
    r["design_fut"] = design_status(r, r["hosts_fut"], r["fut_vcpu"], r["fut_ram"])

    # Performance sizing (Live Optics) - GHz demand, but never below the RAM requirement
    r["perf_hosts_ghz"], r["perf_hosts_rec"], r["perf_constraint"] = 0, 0, ""
    if r["has_perf"]:
        node_ghz = host_cores * p["clock"] * (1 - p["cpu_buffer"] / 100) * PERF_TARGET_UTIL
        r["perf_hosts_ghz"] = ceil_div(r["perf_ghz"], node_ghz)
        perf_raw = max(r["perf_hosts_ghz"], r["hosts_ram"])
        r["perf_constraint"] = "CPU (GHz)" if r["perf_hosts_ghz"] >= r["hosts_ram"] else "RAM"
        r["perf_hosts_rec"] = max(perf_raw + ha, min_hosts)

    # Licensing
    r["fut_lic_cores"] = r["hosts_fut"] * license_cores_per_host(p["sockets"], p["cores"])
    r["lic_diff"] = r["fut_lic_cores"] - r["cur_lic_cores"]
    return r


def sizing_logic_lines(r):
    """Plain-text explanation, shared by the UI and the report."""
    lines = [
        f"Workload: {r['tot_vcpu']:,.0f} vCPU and {r['tot_ram']:,.0f} GB vRAM across {r['tot_vms']:,} VMs.",
        f"Effective host after overhead: {r['eff_cores']:.1f} cores (x {r['vcpu_ratio']:.1f}:1 = "
        f"{r['eff_cores'] * r['vcpu_ratio']:,.0f} vCPU) and {r['eff_ram']:,.0f} GB RAM.",
        f"Hosts needed: CPU {r['hosts_cpu']}, RAM {r['hosts_ram']} -> {r['constraint']} bound, {r['raw_hosts']} active hosts.",
        f"Final: {r['raw_hosts']} + {r['ha_nodes']} HA = {r['raw_hosts'] + r['ha_nodes']}"
        + (f", raised to the {r['min_hosts']}-host minimum cluster size." if r["min_applied_now"] else "."),
        f"Hypervisor: {r['hypervisor']} - {r['hv_ratio']:.1f}:1 vCPU:pCPU is the {r['hv_kind']} guidance. {r['hv_note']}",
    ]
    if r["override_now"]:
        lines.append(f"Manual override: {r['hosts_now']} host{'s' if r['hosts_now'] != 1 else ''} used instead of the calculated {r['calc_hosts_now']}"
                     + (f" (future: calculated {r['calc_hosts_fut']})." if r["override_fut"] else "."))
    return lines


def host_label(r, which):
    """Headline text for the Day 1 / future host count."""
    if which == "now":
        hosts, calc, forced = r["hosts_now"], r["calc_hosts_now"], r["override_now"]
    else:
        hosts, calc, forced = r["hosts_fut"], r["calc_hosts_fut"], r["override_fut"]
    return (f"{hosts} Hosts", f"Manual override (calculated: {calc})") if forced else (f"{hosts} Hosts", "")


# =============================================================================
# 6. HTML REPORT
# =============================================================================
REPORT_CSS = """
:root {
  --ahead-blue:#009FDC; --bright-blue:#1C4CBF; --impact-green:#00B388; --navy:#123A61; --deep-navy:#102033;
  --light-blue:#67C3E9; --light-teal:#6ECEB2; --deep-teal:#006666; --surface:#EBF0F2; --white:#FFFFFF;
  --charcoal:#2A2A2A; --gray:#333333; --footer-gray:#8C9BAB;
  --font:"Poppins", Verdana, sans-serif; --serif:"PT Serif", Georgia, serif;
}
* { box-sizing: border-box; }
html { scroll-behavior: smooth; }
body { margin: 0; font-family: "Poppins", Verdana, sans-serif; font-weight: 400; color: var(--charcoal); background: var(--white); -webkit-font-smoothing: antialiased; }
.wrap { max-width: 1140px; margin: 0 auto; padding: 0 32px; }
a { color: inherit; }

/* Top navigation */
.topnav { position: sticky; top: 0; z-index: 20; background: var(--deep-navy); border-bottom: 1px solid rgba(255,255,255,.08); }
.topnav .wrap { display: flex; align-items: center; gap: 28px; height: 64px; }
.topnav .brand img { height: 18px; display: block; }
.topnav a.link { color: rgba(255,255,255,.85); text-decoration: none; font-size: 14px; }
.topnav a.link:hover { color: var(--white); }
.btn { margin-left: auto; background: var(--bright-blue); color: var(--white); border: 0; border-radius: 9999px; padding: 10px 22px; font: 500 13px var(--font); cursor: pointer; }

/* Hero */
.hero { position: relative; overflow: hidden; color: var(--white); padding: 72px 0 36px;
        background: linear-gradient(69deg, #123A61 0%, #102033 100%); }
.hero::after { content: ""; position: absolute; right: -180px; top: -140px; width: 560px; height: 560px; border-radius: 50%;
               border: 2px solid rgba(103,195,233,.18); pointer-events: none; }
.prepared { display: flex; align-items: center; flex-wrap: wrap; gap: 18px; margin-bottom: 30px; }
.eyebrow { font-size: 12px; letter-spacing: .28em; text-transform: uppercase; font-weight: 500; }
.hero .eyebrow.muted { color: rgba(255,255,255,.72); }
.hero .eyebrow.kicker { color: var(--light-blue); }
.chip { background: var(--white); color: var(--navy); border-radius: 10px; padding: 12px 22px; font-weight: 600; font-size: 18px; min-height: 52px; display: flex; align-items: center; }
.chip img { max-height: 36px; max-width: 220px; display: block; }
.divider { width: 1px; height: 26px; background: rgba(255,255,255,.25); }
.hero h1 { font-size: 56px; line-height: 1.12; font-weight: 500; color: var(--light-blue); margin: 18px 0 24px; max-width: 980px; }
.hero h1 em { font-family: var(--serif); font-style: italic; font-weight: 400; color: var(--light-teal); }
.hero .lede { font-size: 19px; line-height: 1.65; max-width: 740px; color: rgba(255,255,255,.88); margin: 0; }
.stats { display: grid; grid-template-columns: repeat(4, 1fr); margin-top: 52px; padding-top: 28px; border-top: 1px solid rgba(255,255,255,.15); }
.stat { padding: 0 24px; border-left: 1px solid rgba(255,255,255,.15); }
.stat:first-child { padding-left: 0; border-left: 0; }
.stat .num { font-size: 36px; font-weight: 500; color: var(--light-blue); line-height: 1.2; }
.stat .cap { font-size: 14px; line-height: 1.6; color: rgba(255,255,255,.75); margin-top: 6px; }
.meta { margin-top: 30px; font-size: 12px; color: rgba(255,255,255,.7); display: flex; align-items: center; gap: 10px; }
.dot { width: 8px; height: 8px; border-radius: 50%; background: var(--impact-green); display: inline-block; }

/* Sections */
section { padding: 80px 0; }
section.alt { background: var(--surface); }
section .eyebrow { color: var(--ahead-blue); }
h2 { font-size: 34px; letter-spacing: .06em; text-transform: uppercase; color: var(--navy); font-weight: 500; margin: 10px 0 0; }
h2::after { content: ""; display: block; width: 100px; height: 4px; background: var(--navy); margin-top: 22px; }
.section-lede { font-size: 19px; line-height: 1.65; color: var(--navy); max-width: 880px; margin: 28px 0 36px; }
h3 { font-size: 13px; letter-spacing: .22em; text-transform: uppercase; color: var(--navy); font-weight: 600; margin: 40px 0 16px; }

/* Cards */
.cards { display: grid; gap: 20px; grid-template-columns: repeat(auto-fit, minmax(240px, 1fr)); }
.cards.two { grid-template-columns: repeat(auto-fit, minmax(420px, 1fr)); }
.card { background: var(--white); border-radius: 16px; padding: 26px; box-shadow: 0 6px 22px rgba(16,32,51,.08);
        border-top: 4px solid var(--ahead-blue); break-inside: avoid; page-break-inside: avoid; }
.card.green { border-top-color: var(--impact-green); }
.card.navy { border-top-color: var(--navy); }
.card.lead { background: linear-gradient(69deg, rgba(0,179,136,.08) 0%, rgba(0,179,136,0) 60%), #FFFFFF; }
.card .label { font-size: 11px; letter-spacing: .22em; text-transform: uppercase; color: var(--bright-blue); font-weight: 500; margin-bottom: 12px; }
.card .big { font-size: 34px; font-weight: 600; color: var(--navy); line-height: 1.15; }
.card .big small { font-size: 16px; font-weight: 500; color: var(--gray); }
.card p { font-size: 14px; line-height: 1.6; color: var(--gray); margin: 12px 0 0; }
.card p strong { color: var(--charcoal); }
.kpi .big { font-size: 28px; }
.kv { width: 100%; border-collapse: collapse; font-size: 14px; }
.kv td { padding: 9px 0; border-bottom: 1px solid var(--surface); vertical-align: top; }
.kv tr:last-child td { border-bottom: 0; }
.kv td:first-child { color: var(--gray); width: 46%; padding-right: 12px; }
.kv td:last-child { color: var(--navy); font-weight: 500; }

/* Status pills */
.pill { display: inline-flex; align-items: center; gap: 7px; border-radius: 9999px; padding: 5px 12px; font-size: 12px; font-weight: 600; margin-top: 14px; }
.pill::before { content: ""; width: 7px; height: 7px; border-radius: 50%; background: currentColor; }
.pill.ok { background: rgba(0,179,136,.14); color: var(--deep-teal); }
.pill.warn { background: var(--navy); color: var(--white); }
.pill.warn::before { content: "\\26A0"; width: auto; height: auto; border-radius: 0; background: none; font-size: 12px; }
.pill.info { background: rgba(0,159,220,.12); color: var(--bright-blue); }
.checks { list-style: none; padding: 0; margin: 14px 0 0; font-size: 13px; line-height: 1.55; }
.checks li { padding-left: 22px; position: relative; margin-bottom: 6px; color: var(--gray); }
.checks li::before { content: ""; position: absolute; left: 4px; top: 7px; width: 8px; height: 8px; border-radius: 50%; background: var(--impact-green); }
.checks li.bad { color: var(--navy); font-weight: 600; }
.checks li.bad::before { content: "\\26A0"; background: none; width: auto; height: auto; top: 0; left: 0; color: var(--navy); }
.tag { display: inline-block; margin-top: 10px; font-size: 12px; font-weight: 600; color: var(--bright-blue); background: rgba(0,159,220,.12); border-radius: 9999px; padding: 4px 12px; }

/* Navy design-check panel */
.panel { background: var(--navy); border-radius: 22px; padding: 30px; color: var(--white); margin-top: 28px; break-inside: avoid; page-break-inside: avoid; }
.panel-head { display: flex; justify-content: space-between; gap: 24px; flex-wrap: wrap; align-items: baseline; margin-bottom: 22px; }
.panel-head .eyebrow { color: var(--light-blue); }
.panel-head p { margin: 0; font-size: 14px; color: rgba(255,255,255,.78); max-width: 620px; }
.pgrid { display: grid; gap: 16px; grid-template-columns: repeat(auto-fit, minmax(220px, 1fr)); }
.pcard { background: rgba(255,255,255,.06); border: 1px solid rgba(255,255,255,.12); border-radius: 14px; padding: 20px; }
.pcard .eyebrow { font-size: 11px; color: rgba(255,255,255,.72); letter-spacing: .2em; }
.pcard h4 { font-size: 22px; font-weight: 600; margin: 10px 0 8px; color: var(--white); }
.pcard p { font-size: 13.5px; line-height: 1.55; color: rgba(255,255,255,.78); margin: 0; }
.panel .pill.ok { background: rgba(0,179,136,.2); color: var(--light-teal); }
.panel .pill.warn { background: var(--white); color: var(--navy); }
.panel .pill.info { background: rgba(103,195,233,.18); color: var(--light-blue); }

/* Tables */
.table-card { background: var(--white); border-radius: 16px; overflow: hidden; box-shadow: 0 6px 22px rgba(16,32,51,.08); margin-top: 28px; break-inside: avoid; }
.compare { width: 100%; border-collapse: collapse; font-size: 14px; }
.compare th { background: var(--navy); color: var(--white); font-size: 11px; letter-spacing: .16em; text-transform: uppercase; font-weight: 500; text-align: left; padding: 14px 18px; }
.compare td { padding: 12px 18px; border-bottom: 1px solid var(--surface); color: var(--gray); }
.compare tr:last-child td { border-bottom: 0; }
.compare td:first-child { color: var(--navy); font-weight: 500; }
.compare td.hl { background: rgba(0,159,220,.07); color: var(--navy); font-weight: 600; }
.note { background: rgba(0,159,220,.07); border: 1px solid rgba(0,159,220,.25); border-radius: 12px; padding: 14px 18px; font-size: 14px; line-height: 1.6; color: var(--navy); margin-top: 24px; }
.logic { margin: 0; padding-left: 20px; font-size: 14px; line-height: 1.75; color: var(--gray); }

/* Ratio above hypervisor guidance */
.over { color: #C62828 !important; font-weight: 700 !important; }
.hero .over, .panel .over { color: #FF8A80 !important; }
.compare td.hl.over { color: #C62828; }
.pill.hv { background: rgba(198,40,40,.12); color: #C62828; }
.panel .pill.hv { background: rgba(255,138,128,.2); color: #FF8A80; }
.hv-note { padding: 12px 18px; font-size: 12.5px; color: var(--gray); border-top: 1px solid var(--surface); }

/* Footer */
footer { background: var(--deep-navy); color: rgba(255,255,255,.7); padding: 34px 0; font-size: 12px; }
footer .wrap { display: flex; align-items: center; justify-content: space-between; gap: 20px; flex-wrap: wrap; }
footer img { height: 16px; display: block; }

@media (max-width: 860px) {
  .hero h1 { font-size: 38px; }
  .stats { grid-template-columns: repeat(2, 1fr); row-gap: 24px; }
  .stat:nth-child(3) { padding-left: 0; border-left: 0; }
  .topnav a.link { display: none; }
  .cards.two { grid-template-columns: 1fr; }
}
@media print {
  * { -webkit-print-color-adjust: exact; print-color-adjust: exact; }
  .topnav { position: static; }
  .btn { display: none; }
  section { padding: 40px 0; }
  .hero { padding: 40px 0 28px; }
}
"""


def numa_messages(r):
    """Returns list of (is_problem, text) - shared by UI and report."""
    if r["tot_vms"] == 0:
        return [(False, "No VM data found.")]
    msgs = []
    if r["wide_cpu_count"]:
        msgs.append((True, f"Wide CPU: {r['wide_cpu_count']} VM(s) exceed {r['tgt_numa_cores']} cores per socket. "
                           f"Largest: '{r['name_max_cpu']}' ({r['max_vm_cpu']} vCPU)."))
    if r["wide_ram_count"]:
        msgs.append((True, f"Wide RAM: {r['wide_ram_count']} VM(s) exceed {r['tgt_numa_ram']:,.0f} GB per socket. "
                           f"Largest: '{r['name_max_ram']}' ({r['max_vm_ram']:,.0f} GB)."))
    return msgs or [(False, "Healthy: all VMs fit within the target NUMA boundaries.")]


def perf_insight(r):
    """Returns (is_consolidation, title, message) or None."""
    if not r["has_perf"]:
        return None
    if r["perf_hosts_rec"] < r["hosts_now"]:
        return (True, "Consolidation Opportunity",
                f"Measured workload ({r['lo_basis']}) fits on {r['perf_hosts_rec']} hosts "
                f"({r['perf_constraint']} bound) vs {r['hosts_now']} {'forced' if r['override_now'] else 'by allocation'}.")
    if r["perf_hosts_rec"] == r["hosts_now"]:
        return (True, "Aligned",
                f"Measured workload ({r['lo_basis']}) also needs {r['perf_hosts_rec']} hosts "
                f"({r['perf_constraint']} bound), matching the {'forced' if r['override_now'] else 'allocation-based'} host count.")
    return (False, "Performance Risk",
            f"Measured workload ({r['lo_basis']}) needs {r['perf_hosts_rec']} hosts "
            f"({r['perf_constraint']} bound), more than the {'forced' if r['override_now'] else 'allocation-based'} count of {r['hosts_now']}. "
            f"The environment may be running hot.")


def _pct(x):
    return f"{x:.0%}" if x is not None else "n/a"


def current_env(r):
    """Current-state facts. Shared by the UI and the HTML report so both always match.
    Returns (headline_metrics, groups) where groups = [(title, [(label, value, is_warning)])]."""
    if r["vm_on"] is None:
        vm_line = f"{r['tot_vms']:,}"
    elif r["include_off"]:
        vm_line = f"{r['tot_vms']:,} ({r['vm_on']:,} on / {r['vm_off']:,} off)"
    else:
        vm_line = f"{r['tot_vms']:,} powered on ({r['vm_off']:,} powered-off excluded)"

    mem_alloc_warn = r["cur_mem_alloc"] is not None and r["cur_mem_alloc"] > 1
    headline = [
        ("Hosts", f"{r['cur_host_count']:,}", False),
        ("Physical Cores", f"{r['cur_cores']:,.0f}", False),
        ("Host RAM", f"{r['cur_total_ram_gb']:,.0f} GB", False),
        ("vCPU:pCPU", fmt_ratio(r["cur_ratio"]), over_hv(r, r["cur_ratio"])),
        ("VMs", f"{r['tot_vms']:,}", False),
        ("vCPU", f"{r['tot_vcpu']:,.0f}", False),
        ("vRAM", f"{r['tot_ram']:,.0f} GB", False),
        ("vRAM / Host RAM", _pct(r["cur_mem_alloc"]), False),
    ]

    cores_per_host = r["cur_cores"] / r["cur_host_count"] if r["cur_host_count"] else 0
    speed = (f"{r['cur_ghz']:.2f} GHz ({r['cur_ghz_total']:,.0f} GHz cluster total)"
             if r["cur_ghz"] else "Not in source file")
    hosts = [
        ("Hosts in scope", f"{r['cur_host_count']:,}", False),
        ("Typical host", f"{r['cur_sockets']} sockets x {r['cur_numa_cores']:.0f} cores, "
                         f"{r['cur_ram_per_host']:,.0f} GB RAM", False),
        ("Total physical cores", f"{r['cur_cores']:,.0f} (avg {cores_per_host:.0f} per host)", False),
        ("Total RAM", f"{r['cur_total_ram_gb']:,.0f} GB", False),
        ("CPU model", r["cur_cpu_model"], False),
        ("CPU speed", speed, False),
        ("NUMA node (per socket)", f"{r['cur_numa_cores']:.0f} cores | {r['cur_numa_ram']:,.0f} GB", False),
        ("ESXi version", r["esxi_version"], False),
    ]
    vms = [
        ("VMs in scope", vm_line, False),
        ("Total vCPU", f"{r['tot_vcpu']:,.0f}", False),
        ("Total vRAM", f"{r['tot_ram']:,.0f} GB", False),
        ("Average VM", f"{r['avg_vcpu']:.1f} vCPU | {r['avg_ram']:,.1f} GB", False),
        ("Largest VM (CPU)", f"{r['name_max_cpu']} ({r['max_vm_cpu']} vCPU)", False),
        ("Largest VM (RAM)", f"{r['name_max_ram']} ({r['max_vm_ram']:,.0f} GB)", False),
        ("VMs per host", f"{r['cur_vms_per_host']:.1f}" if r["cur_vms_per_host"] else "n/a", False),
    ]
    density = [
        ("vCPU:pCPU (all hosts)", fmt_ratio(r["cur_ratio"]),
         "red" if over_hv(r, r["cur_ratio"]) else r["cur_ratio"] > r["vcpu_ratio"]),
        ("vCPU:pCPU (1 host failed)", fmt_ratio(r["cur_ratio_n1"]),
         "red" if over_hv(r, r["cur_ratio_n1"]) else
         (r["cur_ratio_n1"] is not None and r["cur_ratio_n1"] > r["vcpu_ratio"])),
        (f"{r['hypervisor']} guidance", f"{r['hv_ratio']:.1f}:1 vCPU:pCPU ({r['hv_kind']})", False),
        ("vRAM allocated vs host RAM", _pct(r["cur_mem_alloc"]), mem_alloc_warn),
        ("vRAM vs host RAM (1 host failed)", _pct(r["cur_mem_alloc_n1"]),
         r["cur_mem_alloc_n1"] is not None and r["cur_mem_alloc_n1"] > 1),
    ]
    if r["has_perf"]:
        density.append((f"Measured CPU ({r['lo_basis']})", f"{r['perf_ghz']:,.1f} GHz", False))
        density.append(("Measured CPU utilization", _pct(r["cur_cpu_util"]) if r["cur_cpu_util"] is not None
                        else "n/a (host CPU speed not in file)", False))
    density += [
        ("Licensed cores (current)", f"{r['cur_lic_cores']:,}", False),
        ("License edition", r["lic_edition"], False),
    ]
    sto = r["storage"]
    storage = [
        ("VMDK provisioned", f"{r['vm_prov_tb']:,.1f} TB", False),
        ("VMDK in use", f"{r['vm_used_tb']:,.1f} TB", False),
        ("Guest OS used (backup scope)", f"{r['guest_used_tb']:,.1f} TB", False),
    ]
    if r["rdm_cnt"]:
        storage.append(("RDM disks", f"{r['rdm_cnt']} ({r['rdm_tb']:,.1f} TB)", False))
    if sto["cap"] > 0:
        storage += [
            ("Datastore capacity", f"{sto['cap']:,.1f} TB", False),
            ("Datastore used / free", f"{sto['used']:,.1f} TB / {sto['free']:,.1f} TB "
                                      f"({_pct(sto['used'] / sto['cap'])} used)", False),
        ]
    else:
        storage.append(("Datastores", "Not in source file", False))
    storage.append(("Storage scope", sto["note"], False))

    groups = [("Host Inventory", hosts), ("VM Workload", vms),
              ("Density & Utilization", density), ("Storage", storage)]
    return headline, groups


def comparison_rows(r):
    """Current vs recommended vs future - shared by UI and report."""
    lic_per_host = license_cores_per_host(r["sockets"], r["cores"])
    cur_cores, now_cores, fut_cores = r["cur_cores"], r["hosts_now"] * r["host_cores"], r["hosts_fut"] * r["host_cores"]
    now_ram, fut_ram = r["hosts_now"] * r["ram"], r["hosts_fut"] * r["ram"]
    return [
        ("Hosts", f"{r['cur_host_count']:,}", f"{r['hosts_now']:,} (N+{r['ha_nodes']})", f"{r['hosts_fut']:,} (N+{r['ha_nodes']})"),
        ("Physical cores", f"{cur_cores:,.0f}", f"{now_cores:,}", f"{fut_cores:,}"),
        ("Host RAM", f"{r['cur_total_ram_gb']:,.0f} GB", f"{now_ram:,.0f} GB", f"{fut_ram:,.0f} GB"),
        ("Workload", f"{r['tot_vcpu']:,.0f} vCPU | {r['tot_ram']:,.0f} GB",
         f"{r['tot_vcpu']:,.0f} vCPU | {r['tot_ram']:,.0f} GB", f"{r['fut_vcpu']:,.0f} vCPU | {r['fut_ram']:,.0f} GB"),
        ("vCPU:pCPU (all hosts)", fmt_ratio(r["cur_ratio"]), fmt_ratio(r["ratio_now_all"]), fmt_ratio(r["ratio_fut_all"])),
        ("vCPU:pCPU (HA failover)", fmt_ratio(r["cur_ratio_n1"]), fmt_ratio(r["ratio_now_failover"]), fmt_ratio(r["ratio_fut_failover"])),
        ("vRAM / host RAM", _pct(r["cur_mem_alloc"]),
         _pct(r["tot_ram"] / now_ram if now_ram else None), _pct(r["fut_ram"] / fut_ram if fut_ram else None)),
        ("Licensed cores", f"{r['cur_lic_cores']:,}", f"{r['hosts_now'] * lic_per_host:,}", f"{r['fut_lic_cores']:,}"),
    ]


def file_data_uri(path, mime="image/png"):
    if not path or not os.path.exists(path):
        return None
    with open(path, "rb") as f:
        return f"data:{mime};base64,{base64.b64encode(f.read()).decode()}"


def comparison_flags(r):
    """(row, col) cells of comparison_rows() whose ratio is above hypervisor guidance. col 1=current, 2=day1, 3=future."""
    cells = {(4, 1): r["cur_ratio"], (4, 2): r["ratio_now_all"], (4, 3): r["ratio_fut_all"],
             (5, 1): r["cur_ratio_n1"], (5, 2): r["ratio_now_failover"], (5, 3): r["ratio_fut_failover"]}
    return {rc for rc, v in cells.items() if over_hv(r, v)}


def generate_html_report(r, customer_logo_uri=None):
    """AHEAD-branded, self-contained HTML report (fonts load from Google Fonts with a Verdana fallback)."""
    e = lambda v: escape(str(v))  # noqa: E731 - every dynamic string is escaped
    now = datetime.now()
    headline, groups = current_env(r)
    day1_col = "Proposed (Day 1)" if r["override_now"] else "Recommended (Day 1)"
    ahead_white = file_data_uri(WHITE_LOGO_PATH)
    plural = lambda n, w: f"{n:,} {w}{'' if n == 1 else 's'}"  # noqa: E731

    def pill(ok, text, kind=None):
        cls = kind or ("ok" if ok else "warn")
        return f"<span class='pill {cls}'>{e(text)}</span>"

    def kv_table(rows):
        cells = []
        for k, v, warn in rows:
            if warn == "red":
                val = f"<strong class='over'>&#9888; {e(v)}</strong>"
            else:
                val = f"<strong>&#9888; {e(v)}</strong>" if warn else e(v)
            cells.append(f"<tr><td>{e(k)}</td><td>{val}</td></tr>")
        return "<table class='kv'>" + "".join(cells) + "</table>"

    def ratio(x):
        """Ratio text, bold red when above hypervisor guidance."""
        return f"<strong class='over'>{fmt_ratio(x)}</strong>" if over_hv(r, x) else fmt_ratio(x)

    flags = comparison_flags(r)

    def over_cls(row, col, inline=False):
        if (row, col) not in flags:
            return ""
        return " over" if inline else " class='over'"

    def checks_list(checks):
        return "<ul class='checks'>" + "".join(
            f"<li class='{'bad' if bad else ''}'>{e(t)}</li>" for bad, t in checks) + "</ul>"

    # ---------- Hero ----------
    prepared = (f"<div class='chip'><img src='{customer_logo_uri}' alt='{e(r['customer'])}'></div>"
                if customer_logo_uri else f"<div class='chip'>{e(r['customer'])}</div>")
    constraint_txt = "CPU and RAM balanced" if r["constraint"] == "Balanced" else f"{r['constraint']} bound"
    day1_cap = (f"Proposed Day 1 · manual override (calculated {r['calc_hosts_now']})" if r["override_now"]
                else f"Recommended Day 1 · N+{r['ha_nodes']} · {constraint_txt}")
    stats = [
        (plural(r["hosts_now"], "host"), day1_cap),
        (plural(r["hosts_fut"], "host"), f"After {r['years']} years at {r['growth']*100:.0f}% annual growth"
         + (" · override" if r["override_fut"] else "")),
        (f"{r['tot_vcpu']:,.0f} vCPU", f"{r['tot_vms']:,} VMs · {r['tot_ram']:,.0f} GB vRAM in scope today"),
        (fmt_ratio(r["ratio_now_failover"]), f"vCPU:pCPU during HA failover · {r['hypervisor']} "
         f"{r['hv_kind']} {r['hv_ratio']:.1f}:1 · design limit {r['vcpu_ratio']:.1f}:1"),
    ]
    stats_html = "".join(
        f"<div class='stat'><div class='num{' over' if i == 3 and over_hv(r, r['ratio_now_failover']) else ''}'>{e(n)}</div>"
        f"<div class='cap'>{e(c)}</div></div>" for i, (n, c) in enumerate(stats))
    hero_lede = (f"Host sizing for {r['scope_label']} from the {r['source_type']} export. Today "
                 f"{plural(r['cur_host_count'], 'host')} with {r['cur_cores']:,.0f} cores and "
                 f"{r['cur_total_ram_gb']:,.0f} GB RAM run {plural(r['tot_vms'], 'VM')}. "
                 f"The target is {plural(r['hosts_now'], 'node')} of {r['sockets']} x {r['cores']}-core, "
                 f"{r['ram']:,} GB hosts, growing to {r['hosts_fut']} over {r['years']} years.")

    # ---------- Section 1: current environment ----------
    kpi_html = "".join(
        f"<div class='card kpi {'navy' if i < 4 else ''}'><div class='label'>{e(k)}</div>"
        f"<div class='big{' over' if flag else ''}'>{e(v)}</div>"
        + (f"<p class='over'>Above {e(hv_label(r))}</p>" if flag else "") + "</div>"
        for i, (k, v, flag) in enumerate(headline))
    groups_html = "".join(
        f"<div class='card {'green' if i % 2 else ''}'><div class='label'>{e(title)}</div>{kv_table(rows)}</div>"
        for i, (title, rows) in enumerate(groups))
    cur_lede = (f"{plural(r['cur_host_count'], 'host')} with {r['cur_cores']:,.0f} physical cores and "
                f"{r['cur_total_ram_gb']:,.0f} GB RAM run {plural(r['tot_vms'], 'VM')} "
                f"({r['tot_vcpu']:,.0f} vCPU, {r['tot_ram']:,.0f} GB vRAM) at {fmt_ratio(r['cur_ratio'])} vCPU:pCPU"
                + (f" (above the {hv_label(r)})" if over_hv(r, r["cur_ratio"]) else "")
                + (f", with vRAM at {_pct(r['cur_mem_alloc'])} of host RAM." if r["cur_mem_alloc"] is not None else "."))

    # ---------- Section 2: recommendation ----------
    def host_card(which):
        now = which == "now"
        hosts = r["hosts_now"] if now else r["hosts_fut"]
        title = "Current refresh requirement" if now else f"Future requirement · {r['years']} years"
        tag = host_label(r, which)[1]
        all_r, fo_r = (r["ratio_now_all"], r["ratio_now_failover"]) if now else (r["ratio_fut_all"], r["ratio_fut_failover"])
        cons = r["constraint"] if now else r["fut_constraint"]
        detail = (f"<strong>N+{r['ha_nodes']} · {e(cons)} constraint.</strong> "
                  + (f"Sized for {r['tot_vcpu']:,.0f} vCPU and {r['tot_ram']:,.0f} GB vRAM."
                     if now else f"Projected {r['fut_vcpu']:,.0f} vCPU and {r['fut_ram']:,.0f} GB vRAM at "
                                 f"{r['growth']*100:.0f}% annual growth.")
                  + f" vCPU:pCPU {ratio(all_r)} across all hosts, {ratio(fo_r)} during HA failover"
                  + (f" <span class='over'>(above the {e(hv_label(r))})</span>." if over_hv(r, all_r) or over_hv(r, fo_r)
                     else f" ({e(hv_label(r))})."))
        return (f"<div class='card {'green lead' if now else ''}'><div class='label'>{e(title)}</div>"
                f"<div class='big'>{e(plural(hosts, 'host'))}</div>"
                + (f"<span class='tag'>{e(tag)}</span>" if tag else "")
                + f"<p>{detail}</p>{checks_list(r['checks_now' if now else 'checks_fut'])}</div>")

    d = r["design_now"]
    panel_cards = [
        ("Target node", f"{r['host_cores']} cores · {r['ram']:,} GB",
         f"{r['sockets']} socket{'' if r['sockets'] == 1 else 's'} x {r['cores']} cores. {r['cpu_buffer']}% CPU and {r['ram_buffer']}% RAM reserved for overhead.",
         pill(True, "Design input", "info")),
        ("CPU", f"{d['cpu_ratio']:.1f}:1 at failover",
         f"vCPU per usable core with {plural(r['ha_nodes'], 'host')} down. Design limit {r['vcpu_ratio']:.1f}:1; "
         f"{r['hypervisor']} {r['hv_kind']} {r['hv_ratio']:.1f}:1 (vCPU:pCPU at failover {fmt_ratio(r['ratio_now_failover'])}).",
         pill(d["cpu_ok"], "Within limit" if d["cpu_ok"] else "Exceeds limit")
         + (f" <span class='pill hv'>Above {e(hv_label(r))}</span>" if over_hv(r, r["ratio_now_failover"]) else "")),
        ("Memory", f"{d['ram_use']:.0%} of usable RAM",
         f"{r['tot_ram']:,.0f} GB vRAM against {d['usable_ram']:,.0f} GB usable during HA failover.",
         pill(d["ram_ok"], "Fits" if d["ram_ok"] else "Short on RAM")),
        ("Resilience", f"N+{r['ha_nodes']} · {plural(r['hosts_now'], 'host')}",
         f"Minimum cluster size {r['min_hosts']}. {plural(max(d['active'], 0), 'host')} carry the workload after a failure.",
         pill(d["ha_ok"] and d["min_ok"], "Met" if d["ha_ok"] and d["min_ok"] else "Not met")),
    ]
    panel_html = "".join(
        f"<div class='pcard'><div class='eyebrow'>{e(a)}</div><h4>{e(b)}</h4><p>{e(c)}</p>{pl}</div>"
        for a, b, c, pl in panel_cards)

    compare_html = (f"<table class='compare'><tr><th></th><th>Current</th><th>{e(day1_col)}</th>"
                    f"<th>Future ({r['years']} yrs)</th></tr>"
                    + "".join(
                        f"<tr><td>{e(a)}</td><td{over_cls(i, 1)}>{e(b)}</td><td class='hl{over_cls(i, 2, True)}'>{e(c)}</td>"
                        f"<td{over_cls(i, 3)}>{e(dd)}</td></tr>"
                        for i, (a, b, c, dd) in enumerate(comparison_rows(r))) + "</table>"
                    + f"<div class='hv-note'>Ratios in <span class='over'>bold red</span> are above the "
                      f"{e(hv_label(r))}. {e(r['hv_note'])}</div>")
    logic_html = "<ol class='logic'>" + "".join(f"<li>{e(l)}</li>" for l in sizing_logic_lines(r)) + "</ol>"
    rec_lede = ((f"A manual override sets the cluster at {plural(r['hosts_now'], 'host')}; the calculated "
                 f"requirement is {r['calc_hosts_now']}. ") if r["override_now"] else
                f"{plural(r['hosts_now'], 'host')} meet today's workload with N+{r['ha_nodes']} resilience; the design is "
                f"{constraint_txt}. ") + f"Growth to {r['hosts_fut']} hosts covers the next {r['years']} years."

    # ---------- Section 3: NUMA ----------
    numa = numa_messages(r)
    numa_ok = not any(bad for bad, _ in numa)
    numa_html = checks_list(numa)

    # ---------- Section 4: licensing ----------
    lic_word = "more" if r["lic_diff"] > 0 else "fewer"
    lic_lede = (f"The proposed cluster licenses {r['fut_lic_cores']:,} cores, "
                f"{abs(r['lic_diff']):,} {lic_word} than today's {r['cur_lic_cores']:,}." if r["lic_diff"] else
                f"The proposed cluster licenses the same {r['fut_lic_cores']:,} cores as today.")

    # ---------- Section 5: performance ----------
    perf_html, perf_nav = "", ""
    ins = perf_insight(r)
    if ins:
        good, title, msg = ins
        perf_nav = "<a class='link' href='#performance'>Performance</a>"
        est_note = (f" The 95th percentile was not in the file, so it was estimated as {LO_95TH_FALLBACK:.0%} of peak."
                    if r["perf_estimated"] else "")
        perf_html = f"""
<section id="performance">
  <div class="wrap">
    <div class="eyebrow">Measured, not allocated</div>
    <h2>Performance Sizing</h2>
    <p class="section-lede">{e(msg)}</p>
    <div class="cards">
      <div class="card navy"><div class="label">Allocated</div><div class="big">{r['tot_vcpu']:,.0f} <small>vCPU</small></div><p>Total configured vCPU in scope.</p></div>
      <div class="card"><div class="label">Consumed · {e(r['lo_basis'])}</div><div class="big">{r['perf_ghz']:,.1f} <small>GHz</small></div><p>Aggregate CPU demand from the Live Optics capture.</p></div>
      <div class="card {'green' if good else ''}"><div class="label">Performance-based hosts</div><div class="big">{r['perf_hosts_rec']}</div>
        {pill(good, title)}</div>
    </div>
    <div class="note"><strong>Basis.</strong> Allocation sizing guarantees 100% entitlement. Performance sizing uses the
      {e(r['lo_basis'])} metric at {PERF_TARGET_UTIL:.0%} target CPU utilization and never goes below the RAM requirement.{e(est_note)}</div>
  </div>
</section>"""

    return f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Host Sizing - {e(r['customer'])} | AHEAD</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Poppins:wght@400;500;600&family=PT+Serif:ital@1&display=swap" rel="stylesheet">
<style>{REPORT_CSS}</style>
</head>
<body>
<nav class="topnav"><div class="wrap">
  <a class="brand" href="#top">{f'<img src="{ahead_white}" alt="AHEAD">' if ahead_white else 'AHEAD'}</a>
  <a class="link" href="#environment">Environment</a>
  <a class="link" href="#recommendation">Recommendation</a>
  <a class="link" href="#architecture">NUMA</a>
  <a class="link" href="#licensing">Licensing</a>
  {perf_nav}
  <button class="btn" onclick="window.print()">Print / Save PDF</button>
</div></nav>

<header class="hero" id="top"><div class="wrap">
  <div class="prepared">
    <span class="eyebrow muted">Prepared for</span>{prepared}
    <span class="divider"></span>
    <span class="eyebrow muted">Scope</span><div class="chip">{e(r['scope_label'])}</div>
  </div>
  <div class="eyebrow kicker">{e(r['hypervisor'])} sizing · {e(r['source_type'])} analysis · {now:%b %d, %Y}</div>
  <h1>{e(r['customer'])} host sizing: <em>today</em> + <em>tomorrow</em></h1>
  <p class="lede">{e(hero_lede)}</p>
  <div class="stats">{stats_html}</div>
  <div class="meta"><span class="dot"></span>Generated {now:%b %d, %Y, %I:%M %p} · Source: {e(r['file_name'])} · {e(APP_TITLE)} v{APP_VERSION}</div>
</div></header>

<section id="environment"><div class="wrap">
  <div class="eyebrow">Where you are today</div>
  <h2>Current Environment</h2>
  <p class="section-lede">{e(cur_lede)}</p>
  <div class="cards">{kpi_html}</div>
  <h3>Inventory detail</h3>
  <div class="cards two">{groups_html}</div>
</div></section>

<section class="alt" id="recommendation"><div class="wrap">
  <div class="eyebrow">The short answer</div>
  <h2>Sizing Recommendation</h2>
  <p class="section-lede">{e(rec_lede)}</p>
  <div class="cards two">{host_card("now")}{host_card("fut")}</div>
  <div class="panel">
    <div class="panel-head"><span class="eyebrow">Day 1 design check</span>
      <p>The {plural(r['hosts_now'], 'host')} cluster tested against the CPU ratio, usable RAM and HA rules set for this sizing.</p></div>
    <div class="pgrid">{panel_html}</div>
  </div>
  <div class="table-card">{compare_html}</div>
  <h3>How the number was reached</h3>
  <div class="card navy">{logic_html}</div>
</div></section>

<section id="architecture"><div class="wrap">
  <div class="eyebrow">Will the largest VMs fit</div>
  <h2>Architecture &amp; NUMA</h2>
  <p class="section-lede">Each target socket offers {r['tgt_numa_cores']} cores and {r['tgt_numa_ram']:,.0f} GB RAM.
    VMs larger than one socket span NUMA nodes, which can cost performance.</p>
  <div class="cards">
    <div class="card navy"><div class="label">Current NUMA node</div><div class="big">{r['cur_numa_cores']:.0f} <small>cores</small></div><p>{r['cur_numa_ram']:,.0f} GB per socket today.</p></div>
    <div class="card"><div class="label">Target NUMA node</div><div class="big">{r['tgt_numa_cores']} <small>cores</small></div><p>{r['tgt_numa_ram']:,.0f} GB per socket on the new hosts.</p></div>
    <div class="card {'green' if numa_ok else ''}"><div class="label">Large VM check</div>
      {pill(numa_ok, "All VMs fit" if numa_ok else "Wide VMs found")}{numa_html}</div>
  </div>
</div></section>

<section class="alt" id="licensing"><div class="wrap">
  <div class="eyebrow">Per-core licensing</div>
  <h2>Licensing Impact</h2>
  <p class="section-lede">{e(lic_lede)} A {LIC_MIN_CORES_PER_SOCKET}-core minimum per CPU applies.</p>
  <div class="cards">
    <div class="card navy"><div class="label">Current state</div><div class="big">{r['cur_lic_cores']:,} <small>cores</small></div>
      <p><strong>{e(r['lic_edition'])}</strong><br>{plural(r['cur_host_count'], 'host')} · ESXi {e(r['esxi_version'])}</p></div>
    <div class="card"><div class="label">Future state</div><div class="big">{r['fut_lic_cores']:,} <small>cores</small></div>
      <p>{plural(r['hosts_fut'], 'host')} x {license_cores_per_host(r['sockets'], r['cores'])} licensed cores per host.</p></div>
    <div class="card {'green' if r['lic_diff'] <= 0 else ''}"><div class="label">Net change</div>
      <div class="big">{r['lic_diff']:+,} <small>cores</small></div>
      {pill(r['lic_diff'] <= 0, "Fewer cores to license" if r['lic_diff'] < 0 else ("No change" if r['lic_diff'] == 0 else "More cores to license"), None if r['lic_diff'] <= 0 else "info")}</div>
  </div>
</div></section>
{perf_html}
<footer><div class="wrap">
  {f'<img src="{ahead_white}" alt="AHEAD">' if ahead_white else '<strong>AHEAD</strong>'}
  <span>Confidential - Proprietary to AHEAD · Prepared for {e(r['customer'])}</span>
  <span>Source: {e(r['file_name'])} ({e(r['source_type'])})</span>
</div></footer>
</body>
</html>"""



# =============================================================================
# 7. STREAMLIT UI
# =============================================================================
PRODUCT_NAME = "Host Sizer"
MIN_STREAMLIT = (1, 50)

THEME_MARKER = "# Managed by Host Sizer (sizing_app.py)."
THEME_TOML = THEME_MARKER + """ Delete this line to stop automatic updates.
[theme]
base = "light"
primaryColor = "#1C4CBF"
backgroundColor = "#FFFFFF"
secondaryBackgroundColor = "#EBF0F2"
textColor = "#2A2A2A"
linkColor = "#1C4CBF"
borderColor = "#D6DEE3"
showWidgetBorder = true
baseRadius = "0.6rem"
buttonRadius = "full"
font = "Poppins:https://fonts.googleapis.com/css2?family=Poppins:wght@400;500;600;700&display=swap, Verdana, sans-serif"
headingFontWeights = [700, 600, 600, 600, 600, 600]
dataframeHeaderBackgroundColor = "#EBF0F2"
blueColor = "#009FDC"
greenColor = "#00B388"

[theme.sidebar]
backgroundColor = "#102033"
secondaryBackgroundColor = "#1B2E45"
textColor = "#FFFFFF"
primaryColor = "#67C3E9"
borderColor = "#2B3F58"
"""

APP_CSS = """<style>
.hs-brand { padding: 2px 2px 18px; margin-bottom: 6px; border-bottom: 1px solid rgba(255,255,255,.12); }
.hs-brand .name { font-size: 30px; font-weight: 700; color: #FFFFFF; line-height: 1.05; letter-spacing: -.01em; }
.hs-brand .pb { display: flex; align-items: center; gap: 7px; margin-top: 7px; font-size: 12px; color: rgba(255,255,255,.8); }
.hs-brand .pb img { height: 11px; display: block; }
.hs-label { font-size: 11px; letter-spacing: .14em; text-transform: uppercase; color: #8C9BAB; font-weight: 600; margin: 12px 0 2px; }
.hs-eyebrow { font-size: 11px; letter-spacing: .18em; text-transform: uppercase; color: #1C4CBF; font-weight: 600; margin-bottom: 4px; }
.hs-big { font-size: 32px; font-weight: 700; color: #123A61; line-height: 1.15; }
.hs-big small { font-size: 15px; font-weight: 500; color: #333333; }
.hs-sub { font-size: 13.5px; color: #333333; line-height: 1.55; margin-top: 6px; }
.hs-kv { width: 100%; border-collapse: collapse; font-size: 14px; }
.hs-kv td { padding: 8px 0; border-bottom: 1px solid #EBF0F2; vertical-align: top; }
.hs-kv tr:last-child td { border-bottom: 0; }
.hs-kv td:first-child { color: #333333; width: 48%; padding-right: 12px; }
.hs-kv td:last-child { color: #123A61; font-weight: 500; text-align: right; }
.hs-kv td.warn { color: #123A61; font-weight: 700; }
.hs-kv td.red, .hs-over { color: #C62828 !important; font-weight: 700; }
.hs-cmp { width: 100%; border-collapse: collapse; font-size: 14px; }
.hs-cmp th { background: #123A61; color: #FFFFFF; font-size: 11px; letter-spacing: .14em; text-transform: uppercase;
             font-weight: 500; text-align: left; padding: 11px 14px; }
.hs-cmp td { padding: 10px 14px; border-bottom: 1px solid #EBF0F2; color: #333333; }
.hs-cmp tr:last-child td { border-bottom: 0; }
.hs-cmp td:first-child { color: #123A61; font-weight: 500; }
.hs-cmp td.hs-hl { background: rgba(0,159,220,.07); color: #123A61; font-weight: 600; }
.hs-tile-name { font-size: 15px; font-weight: 600; color: #123A61; word-break: break-all; }
</style>"""

GREETING_PARTS = [(12, "morning"), (17, "afternoon"), (24, "evening")]

DEFAULTS = {
    "sockets": 2, "cores": 24, "ram": 1024, "clock": 2.5, "lo_basis": "95th Percentile",
    "hypervisor": DEFAULT_HYPERVISOR, "vcpu_ratio": HYPERVISORS[DEFAULT_HYPERVISOR][0], "cpu_buffer": 10, "ram_buffer": 10, "min_hosts": 2, "ha_nodes": 1,
    "override": False, "override_hosts": 4, "override_future": True,
    "include_off": True, "growth_pct": 10.0, "years": 3,
    "customer": "My Customer", "scope": [],
    "cmp_view": "Both", "raw_view": "VMs", "raw_cols": "Key columns", "raw_search": "",
}
PARAM_KEYS = ["hypervisor", "sockets", "cores", "ram", "clock", "lo_basis", "vcpu_ratio", "cpu_buffer", "ram_buffer",
              "min_hosts", "ha_nodes", "override", "override_hosts", "override_future",
              "include_off", "years", "customer"]

KEY_COLUMNS = {
    "RVTools": {"VMs": ["VM", "Powerstate", "Cluster", "Host", "CPUs", "Memory", "Provisioned MiB", "In Use MiB",
                        "OS according to the configuration file", "OS according to the VMware Tools"],
                "Hosts": ["Host", "Cluster", "CPU Model", "Speed", "# CPU", "Cores per CPU", "# Cores", "# Memory",
                          "# VMs", "ESX Version"]},
    "LiveOptics": {"VMs": ["VM Name", "Power State", "Cluster", "Host", "Virtual CPU", "Provisioned Memory (MiB)",
                           "Virtual Disk Size (MiB)", "Virtual Disk Used (MiB)", "Guest OS"],
                   "Hosts": ["Host Name", "Cluster", "CPU Model", "CPU Sockets", "CPU Cores", "Memory (KiB)",
                             "CPU Speed (MHz)", "ESX Version", "OS Version"]},
}

# Set once per run in main(); read by the page functions.
CTX = None
PAGES = {}


# ---------- setup helpers ----------
def ensure_theme_config():
    """Write .streamlit/config.toml next to the app (AHEAD theme). Never overwrites a user-managed file.
    Returns True when the file was created or updated (theme applies after a restart)."""
    path = os.path.join(APP_DIR, ".streamlit", "config.toml")
    try:
        if os.path.exists(path):
            with open(path, encoding="utf-8") as f:
                current = f.read()
            if current == THEME_TOML or not current.startswith(THEME_MARKER):
                return False
        os.makedirs(os.path.dirname(path), exist_ok=True)
        with open(path, "w", encoding="utf-8") as f:
            f.write(THEME_TOML)
        return True
    except OSError:
        return False


def init_state():
    ss = st.session_state
    for k, v in DEFAULTS.items():
        ss.setdefault(k, v)
    ss.setdefault("reports", {})      # name -> {"bytes", "uploaded"}
    ss.setdefault("active", None)
    ss.setdefault("cust_logo", None)  # {"bytes", "type", "name"}
    # Re-assigning widget keys stops Streamlit from discarding them when their page isn't shown.
    for k in list(DEFAULTS) + ["active"]:
        if k in ss:
            ss[k] = ss[k]


def first_name():
    try:
        import getpass
        user = getpass.getuser()
    except Exception:  # noqa: BLE001 - any failure just means no name
        return ""
    part = user.replace("_", ".").split(".")[0]
    return part.capitalize() if part.isalpha() else ""


def greeting():
    hour = datetime.now().hour
    word = next(w for h, w in GREETING_PARTS if hour < h)
    name = first_name()
    return f"Good {word}, {name}" if name else f"Good {word}"


def params_from_state():
    ss = st.session_state
    p = {k: ss[k] for k in PARAM_KEYS}
    p["growth"] = ss["growth_pct"] / 100
    p["lo_basis"] = ss["lo_basis"] or DEFAULTS["lo_basis"]   # segmented control can be deselected
    p["hypervisor"] = ss["hypervisor"] or DEFAULT_HYPERVISOR
    return p


def customer_logo_uri():
    logo = st.session_state.get("cust_logo")
    if not logo:
        return None
    return f"data:{logo['type']};base64,{base64.b64encode(logo['bytes']).decode()}"


def report_summary(entry):
    """Quick facts for a Home tile."""
    try:
        source, sheets = entry_sheets(entry)
    except Exception:  # noqa: BLE001
        return None, "Could not read file"
    if source is None:
        return None, "Not an RVTools or Live Optics export"
    vm_sheet, host_sheet = ("vInfo", "vHost") if source == "RVTools" else ("VMs", "ESX Hosts")
    vms = len(drop_templates(sheets[vm_sheet]))
    hosts = len(sheets.get(host_sheet, []))
    return source, f"{hosts:,} hosts · {vms:,} VMs · {len(cluster_list(source, sheets))} clusters"


def build_context():
    """Parse the active report and run the sizing with the current parameters."""
    ss = st.session_state
    name = ss.get("active")
    if not name or name not in ss.reports:
        return None
    entry = ss.reports[name]
    try:
        source, sheets = entry_sheets(entry)
    except Exception as exc:  # noqa: BLE001
        return {"error": f"Could not read '{name}'. Is it a valid .xlsx export?", "exc": exc, "name": name}
    if source is None:
        return {"error": f"'{name}' is not an RVTools export (vInfo tab) or a Live Optics export "
                         f"(VMs + ESX Hosts tabs).", "exc": None, "name": name}
    clusters = cluster_list(source, sheets)
    ss.scope = [c for c in ss.scope if c in clusters]
    scope = ss.scope or None
    scope_label = ALL_CLUSTERS if scope is None else " + ".join(scope)
    p = params_from_state()
    try:
        if source == "RVTools":
            inv = parse_rvtools(sheets, scope, p["include_off"])
        else:
            inv = parse_live_optics(sheets, scope, p["include_off"], p["lo_basis"])
        r = compute_sizing(inv, p)
        r.update(scope_label=scope_label, file_name=entry.get("file_name", name), source_type=source)
        html = generate_html_report(r, customer_logo_uri())
    except KeyError as exc:
        return {"error": f"A required column or tab is missing from '{name}': {exc}", "exc": exc, "name": name}
    except Exception as exc:  # noqa: BLE001
        return {"error": f"Error processing '{name}': {exc}", "exc": exc, "name": name}
    base = os.path.splitext(entry.get("file_name", name))[0]
    safe_scope = "".join(ch if ch.isalnum() or ch in "-_" else "_" for ch in scope_label.replace(" + ", "-"))
    return {"key": name, "name": entry.get("project_name") or name, "source": source, "sheets": sheets, "clusters": clusters, "inv": inv, "r": r,
            "html": html, "file_name": f"{base}_{safe_scope}_Sizing.html", "error": None}


# ---------- shared UI pieces ----------
def kv_html(rows):
    cells = "".join(
        f"<tr><td>{escape(str(k))}</td><td class='{'red' if w == 'red' else ('warn' if w else '')}'>"
        f"{'&#9888; ' if w else ''}{escape(str(v))}</td></tr>"
        for k, v, w in rows)
    return f"<table class='hs-kv'>{cells}</table>"


def eyebrow(text):
    st.html(f"<div class='hs-eyebrow'>{escape(text)}</div>")


def big_value(value, unit="", sub=""):
    unit_html = f" <small>{escape(unit)}</small>" if unit else ""
    sub_html = f"<div class='hs-sub'>{escape(sub)}</div>" if sub else ""
    st.html(f"<div class='hs-big'>{escape(str(value))}{unit_html}</div>{sub_html}")


def ratio_md(r, x):
    """Markdown ratio, bold red when above the hypervisor guidance."""
    return f":red[**{fmt_ratio(x)}**]" if over_hv(r, x) else fmt_ratio(x)


def badge_md(ok, ok_text, bad_text, neutral=False):
    if neutral:
        return f":blue-badge[:material/tune: {ok_text}]"
    return f":green-badge[:material/check_circle: {ok_text}]" if ok else f":orange-badge[:material/warning: {bad_text}]"


def render_checks(checks):
    for bad, text in checks:
        st.markdown((":orange-badge[:material/warning: Shortfall]" if bad else
                     ":green-badge[:material/check_circle: Meets design]") + f" {text}")


def section_title(eyebrow_text, title):
    st.space("medium") if hasattr(st, "space") else st.write("")
    eyebrow(eyebrow_text)
    st.subheader(title, anchor=False)


def page_header(title, subtitle="", ctx=None):
    render_flash()
    left, right = st.columns([3, 1], vertical_alignment="bottom")
    with left:
        st.title(title, anchor=False)
        if subtitle:
            st.caption(subtitle)
        if ctx and not ctx["error"]:
            r = ctx["r"]
            badges = [f":violet-badge[:material/layers: {r['hypervisor']} · {r['hv_ratio']:.1f}:1]",
                      f":blue-badge[:material/description: {ctx['source']}]",
                      f":gray-badge[:material/hub: {r['scope_label']}]"]
            if r["override_now"]:
                badges.append(":orange-badge[:material/lock: Host count override]")
            if not r["include_off"]:
                badges.append(":gray-badge[Powered-off VMs excluded]")
            badges.append(save_status(st.session_state.reports[ctx["key"]]))
            st.markdown(" ".join(badges))
    with right:
        if ctx and not ctx["error"]:
            st.download_button("Download report", ctx["html"], file_name=ctx["file_name"], mime="text/html",
                               type="primary", icon=":material/download:", width="stretch")
            save_popover()
    st.divider()


def context_subtitle(ctx):
    r = ctx["r"]
    return (f"{r['customer']} · {ctx['name']} · {r['cur_host_count']:,} hosts · {r['tot_vms']:,} VMs · "
            f"{r['tot_vcpu']:,.0f} vCPU · {r['tot_ram']:,.0f} GB vRAM")


def scope_bar(ctx):
    if not ctx["clusters"]:
        return
    help_txt = "None selected = all clusters. Pick several clusters to model a consolidation."
    if len(ctx["clusters"]) <= 20:
        st.pills("Cluster scope", ctx["clusters"], selection_mode="multi", key="scope", help=help_txt)
    else:
        st.multiselect("Cluster scope", ctx["clusters"], key="scope", help=help_txt, placeholder="All clusters")
    if len(st.session_state.scope) > 1:
        st.caption(f":material/merge: Consolidation scope: combining {len(st.session_state.scope)} clusters.")


def require_report(title):
    """Shared empty / error state. Returns ctx when the page can render."""
    ctx = CTX
    if ctx is None:
        page_header(title, "No report loaded yet.")
        with st.container(border=True):
            st.markdown("#### Upload a report to get started")
            st.caption("Load an RVTools or Live Optics export on the Home page.")
            st.page_link(PAGES["home"], label="Go to Home", icon=":material/home:")
        return None
    if ctx["error"]:
        page_header(title, ctx["name"])
        st.error(ctx["error"], icon=":material/error:")
        if ctx["exc"] is not None:
            with st.expander("Error details"):
                st.exception(ctx["exc"])
        st.page_link(PAGES["home"], label="Back to Home", icon=":material/home:")
        return None
    return ctx


# ---------- saved projects ----------
SETTING_KEYS = PARAM_KEYS + ["growth_pct", "scope"]


def settings_snapshot():
    ss = st.session_state
    return {k: ss[k] for k in SETTING_KEYS}


def flash(kind, msg):
    """Queue a message from a callback; shown at the top of the next page render."""
    st.session_state._flash = (kind, msg)


def entry_sheets(entry):
    """(source, sheets) for a session report: saved projects carry reduced sheets, uploads carry bytes."""
    if entry.get("sheets") is not None:
        return entry["source"], entry["sheets"]
    return load_workbook(entry["bytes"])


def default_project_name(entry_name):
    entry = st.session_state.reports.get(entry_name, {})
    if entry.get("project_name"):
        return entry["project_name"]
    base = os.path.splitext(entry.get("file_name", entry_name))[0]
    return f"{st.session_state.customer} - {base}"


def _save_project():
    ss = st.session_state
    name = (ss.get("save_name") or "").strip()
    key = ss.get("active")
    if not name or key not in ss.reports:
        flash("error", "Enter a project name before saving.")
        return
    entry = ss.reports[key]
    try:
        source, sheets = entry_sheets(entry)
        reduced = reduce_sheets(source, sheets)
        # Guard: the reduced data must produce exactly the same sizing as the full export.
        p, scope = params_from_state(), (ss.scope or None)
        def run(sh):
            inv = (parse_rvtools(sh, scope, p["include_off"]) if source == "RVTools"
                   else parse_live_optics(sh, scope, p["include_off"], p["lo_basis"]))
            return compute_sizing(inv, p)
        full_r, small_r = run(sheets), run(reduced)
        diffs = [k for k in VERIFY_KEYS if abs((full_r[k] or 0) - (small_r[k] or 0)) > 1e-6]
        if diffs:
            flash("error", f"Not saved: trimming the data changed {', '.join(diffs)}. Please report this.")
            return
        now = datetime.now().isoformat(timespec="seconds")
        path = project_path(name)
        old = entry.get("project_path")
        created = now
        if old and os.path.exists(old):
            try:
                created = read_project(old).get("created", now)
            except Exception:  # noqa: BLE001
                pass
        logo = ss.get("cust_logo")
        payload = {
            "format": PROJECT_FORMAT, "version": PROJECT_VERSION, "app_version": APP_VERSION,
            "name": name, "created": created, "saved": now,
            "settings": settings_snapshot(),
            "customer_logo": ({"type": logo["type"], "name": logo["name"],
                               "b64": base64.b64encode(logo["bytes"]).decode()} if logo else None),
            "source": {"type": source, "file_name": entry.get("file_name", key),
                       "sha256": entry.get("sha256"), "sheets": frames_to_json(reduced)},
            "summary": {"customer": ss.customer, "hypervisor": full_r["hypervisor"],
                        "hosts": full_r["cur_host_count"], "vms": full_r["tot_vms"],
                        "hosts_now": full_r["hosts_now"], "scope": " + ".join(ss.scope) or ALL_CLUSTERS},
        }
        write_project(path, payload)
        if old and old != path and os.path.exists(old):
            os.remove(old)   # renamed: replace the previous file
    except Exception as exc:  # noqa: BLE001
        flash("error", f"Could not save the project: {exc}")
        return
    entry.update(project_path=path, project_name=name, saved_settings=settings_snapshot(),
                 saved_logo=bool(logo), saved_at=now)
    flash("success", f"Saved '{name}' to {path}")


def _open_project(path):
    ss = st.session_state
    try:
        payload = read_project(path)
        sheets = frames_from_json(payload["source"]["sheets"])
    except Exception as exc:  # noqa: BLE001
        flash("error", f"Could not open {os.path.basename(path)}: {exc}")
        return
    name = payload["name"]
    key = f"{name} (project)"
    for k, v in payload.get("settings", {}).items():
        if k in DEFAULTS or k == "growth_pct":
            ss[k] = v
    logo = payload.get("customer_logo")
    ss.cust_logo = ({"type": logo["type"], "name": logo.get("name", "logo"),
                     "bytes": base64.b64decode(logo["b64"])} if logo else None)
    ss.reports[key] = {"sheets": sheets, "source": payload["source"]["type"],
                       "file_name": payload["source"]["file_name"], "sha256": payload["source"].get("sha256"),
                       "uploaded": datetime.now(), "project_path": path, "project_name": name,
                       "saved_settings": settings_snapshot(), "saved_logo": bool(logo), "saved_at": payload.get("saved")}
    ss.active = key
    ss._save_for = None
    ss._goto = "env"


def _delete_project(path):
    try:
        os.remove(path)
    except OSError as exc:
        flash("error", f"Could not delete {os.path.basename(path)}: {exc}")
        return
    for entry in st.session_state.reports.values():
        if entry.get("project_path") == path:
            entry.update(project_path=None, saved_settings=None)
    flash("success", f"Deleted {os.path.basename(path)}")


def save_status(entry):
    if not entry.get("project_path"):
        return ":gray-badge[:material/cloud_off: Not saved]"
    dirty = (entry.get("saved_settings") != settings_snapshot()
             or entry.get("saved_logo") != bool(st.session_state.get("cust_logo")))
    return (":orange-badge[:material/edit: Unsaved changes]" if dirty
            else ":green-badge[:material/check_circle: Saved]")


def save_popover():
    ss = st.session_state
    key = ss.active
    if ss.get("_save_for") != key:
        ss.save_name = default_project_name(key)
        ss._save_for = key
    with st.popover("Save project", icon=":material/save:", width="stretch"):
        st.text_input("Project name", key="save_name")
        target = project_path(ss.save_name or "project")
        current = ss.reports[key].get("project_path")
        if os.path.exists(target) and target != current:
            st.caption(":orange[A different project with this name exists and will be replaced.]")
        st.caption(f"Saves only the columns the sizing uses, plus your settings, to `{target}`.")
        st.button("Save", type="primary", icon=":material/save:", on_click=_save_project, width="stretch")


def render_flash():
    msg = st.session_state.pop("_flash", None)
    if msg:
        kind, text = msg
        (st.success if kind == "success" else st.error)(text, icon=":material/check_circle:" if kind == "success"
                                                         else ":material/error:")


def render_saved_projects():
    projects = list_projects()
    section_title("Saved projects", "Your project library")
    st.caption(f"Stored on this computer in `{PROJECTS_DIR}`. Each project keeps your settings and only the "
               "columns the sizing needs; the rest of the export is discarded.")
    if not projects:
        st.caption("No saved projects yet. Open a report and use **Save project** at the top of any page.")
        return
    for row_start in range(0, len(projects), 3):
        for col, (path, meta) in zip(st.columns(3), projects[row_start:row_start + 3]):
            with col, st.container(border=True):
                if meta.get("error"):
                    st.markdown(":red-badge[:material/error: Unreadable]")
                    st.html(f"<div class='hs-tile-name'>{escape(meta['name'])}</div>")
                    st.caption(meta["error"])
                    continue
                sm = meta.get("summary") or {}
                st.markdown(f":blue-badge[:material/description: {meta['source_type']}] "
                            f":violet-badge[:material/layers: {sm.get('hypervisor', '')}]")
                st.html(f"<div class='hs-tile-name'>{escape(meta['name'])}</div>")
                saved = (meta.get("saved") or "").replace("T", " ")[:16]
                st.caption(f"{sm.get('customer', '')} · {meta['file_name']}  \n"
                           f"{sm.get('hosts', 0):,} hosts · {sm.get('vms', 0):,} VMs · "
                           f"{sm.get('hosts_now', 0)} recommended · {sm.get('scope', '')}  \nSaved {saved}")
                b1, b2 = st.columns(2)
                b1.button("Open", key=f"popen_{path}", icon=":material/folder_open:", type="primary",
                          width="stretch", on_click=_open_project, args=(path,))
                with b2.popover("Delete", icon=":material/delete:", width="stretch"):
                    st.caption(f"Delete **{meta['name']}** permanently from this computer?")
                    st.button("Delete project", key=f"pdel_{path}", type="primary", on_click=_delete_project,
                              args=(path,))


# ---------- sidebar ----------
def render_sidebar():
    sb = st.sidebar
    logo = file_data_uri(WHITE_LOGO_PATH)
    pb = f"Powered by <img src='{logo}' alt='AHEAD'>" if logo else "Powered by AHEAD"
    sb.html(f"<div class='hs-brand'><div class='name'>{escape(PRODUCT_NAME)}</div><div class='pb'>{pb}</div></div>")

    has_report = CTX is not None and not CTX.get("error")
    sb.html("<div class='hs-label'>Workspace</div>")
    sb.page_link(PAGES["home"], label="Home", icon=":material/home:")
    sb.page_link(PAGES["env"], label="Current Environment", icon=":material/dns:", disabled=not has_report)
    sb.page_link(PAGES["rec"], label="Recommendation", icon=":material/insights:", disabled=not has_report)
    sb.page_link(PAGES["report"], label="Report", icon=":material/description:", disabled=not has_report)
    sb.page_link(PAGES["raw"], label="Raw Data", icon=":material/table_view:", disabled=not has_report)

    reports = st.session_state.reports
    if reports:
        sb.html("<div class='hs-label'>Active report</div>")
        sb.selectbox("Active report", list(reports), key="active", label_visibility="collapsed",
                     on_change=lambda: st.session_state.update(scope=[]))
        if has_report:
            sb.caption(f"{CTX['source']} · {CTX['r']['scope_label']}")
    sb.html("<div class='hs-label'>About</div>")
    sb.caption(f"{APP_TITLE} v{APP_VERSION}")


# ---------- pages ----------
def _ingest_uploads():
    ss = st.session_state
    files = ss.get("uploader") or []
    for f in files:
        data = f.getvalue()
        ss.reports[f.name] = {"bytes": data, "file_name": f.name, "uploaded": datetime.now(),
                              "sha256": hashlib.sha256(data).hexdigest()}
    if files:
        ss.active = files[-1].name
        ss.scope = []
        ss._goto = "env"


def _open_report(name):
    st.session_state.active = name
    st.session_state.scope = []
    st.session_state._goto = "env"


def _remove_report(name):
    ss = st.session_state
    ss.reports.pop(name, None)
    if ss.active == name:
        ss.active = next(iter(ss.reports), None)
        ss.scope = []


def page_home():
    n = len(st.session_state.reports)
    page_header(greeting(), f"{n} report{'s' if n != 1 else ''} loaded this session" if n
                else "Upload an RVTools or Live Optics export to start a sizing.")
    with st.container(border=True):
        st.markdown("#### Upload a report")
        st.caption("RVTools (.xlsx with a vInfo tab) or Dell Live Optics (.xlsx with VMs and ESX Hosts tabs). "
                   "Files stay on this computer.")
        st.file_uploader("Upload RVTools or Live Optics Excel file(s)", type=["xlsx"], accept_multiple_files=True,
                         key="uploader", on_change=_ingest_uploads, label_visibility="collapsed")

    if n:
        render_session_reports()
    render_saved_projects()


def render_session_reports():
    section_title("This session", "Jump back in")
    names = list(st.session_state.reports)
    for row_start in range(0, len(names), 3):
        cols = st.columns(3)
        for col, name in zip(cols, names[row_start:row_start + 3]):
            info = st.session_state.reports[name]
            source, summary = report_summary(info)
            with col, st.container(border=True):
                active = name == st.session_state.active
                st.markdown(f":{'blue' if source else 'gray'}-badge[:material/description: {source or 'Unknown'}]"
                            + (" :green-badge[Active]" if active else ""))
                st.html(f"<div class='hs-tile-name'>{escape(name)}</div>")
                st.caption(f"{summary} · {'opened' if info.get('project_path') else 'uploaded'} "
                           f"{info['uploaded']:%I:%M %p}")
                if active:
                    st.markdown(save_status(info))
                elif info.get("project_path"):
                    st.markdown(":green-badge[:material/folder: Saved project]")
                b1, b2 = st.columns(2)
                b1.button("Open", key=f"open_{name}", icon=":material/arrow_forward:", type="primary",
                          width="stretch", disabled=source is None, on_click=_open_report, args=(name,))
                b2.button("Remove", key=f"rm_{name}", icon=":material/delete:", width="stretch",
                          on_click=_remove_report, args=(name,))


def page_environment():
    ctx = require_report("Current Environment")
    if not ctx:
        return
    page_header("Current Environment", context_subtitle(ctx), ctx)
    scope_bar(ctx)
    headline, groups = current_env(ctx["r"])
    r = ctx["r"]
    for row in (headline[:4], headline[4:]):
        for col, (label, value, flag) in zip(st.columns(4), row):
            if flag:
                col.metric(label, value, delta=f"Above {hv_label(r)}", delta_color="inverse", border=True,
                           help=r["hv_note"])
            else:
                col.metric(label, value, border=True)
    section_title("Inventory detail", "Hosts, workload, density and storage")
    for pair in (groups[:2], groups[2:]):
        for col, (title, rows) in zip(st.columns(2), pair):
            with col, st.container(border=True):
                st.markdown(f"#### {title}")
                st.html(kv_html(rows))


def _apply_hv_ratio():
    """Set the design limit to the selected hypervisor's vCPU:pCPU guidance."""
    hv = st.session_state.get("hypervisor") or DEFAULT_HYPERVISOR
    st.session_state.vcpu_ratio = HYPERVISORS[hv][0]


def render_parameters(ctx):
    with st.container(border=True):
        st.markdown("#### Sizing parameters")
        st.caption("Changes apply instantly to every page and to the downloadable report.")
        t0, t1, t2, t3, t4 = st.tabs([":material/layers: Hypervisor", ":material/memory: Target hardware",
                                      ":material/rule: Constraints", ":material/lock: Host count override",
                                      ":material/trending_up: Scope & growth"])
        with t0:
            st.segmented_control("Target hypervisor", list(HYPERVISORS), key="hypervisor", on_change=_apply_hv_ratio,
                                 help="Choosing a hypervisor sets the design limit to its vCPU:pCPU guidance. "
                                      "Ratios above the guidance are flagged in bold red.")
            hv = st.session_state.hypervisor or DEFAULT_HYPERVISOR
            hv_ratio, hv_kind, hv_note = HYPERVISORS[hv]
            c = st.columns([1, 1])
            with c[0]:
                st.slider("Max vCPU:pCPU (design limit)", 1.0, 10.0, step=0.5, key="vcpu_ratio",
                          help="Set automatically from the hypervisor; adjust to size above or below the guidance.")
                st.caption(f"**{hv}: {hv_ratio:.1f}:1 ({hv_kind}).** {hv_note}")
                if abs(st.session_state.vcpu_ratio - hv_ratio) > EPS:
                    above = st.session_state.vcpu_ratio > hv_ratio
                    st.markdown((f":orange-badge[:material/warning: Above {hv} {hv_kind}]" if above else
                                 f":blue-badge[:material/tune: Below {hv} {hv_kind}]")
                                + f" Design limit {st.session_state.vcpu_ratio:.1f}:1 vs guidance {hv_ratio:.1f}:1.")
                    st.button(f"Reset to {hv_ratio:.1f}:1", icon=":material/restart_alt:", on_click=_apply_hv_ratio)
            with c[1]:
                st.dataframe(pd.DataFrame([(k, f"{v[0]:.1f}:1", v[1].capitalize()) for k, v in HYPERVISORS.items()],
                                          columns=["Hypervisor", "vCPU:pCPU", "Type"]),
                             hide_index=True, width="stretch")
        with t1:
            c = st.columns(4)
            c[0].number_input("Sockets / host", 1, 4, key="sockets")
            c[1].number_input("Cores / socket", 4, 128, key="cores")
            c[2].number_input("RAM / host (GB)", 64, 8192, key="ram")
            c[3].number_input("CPU speed (GHz)", 1.0, 5.0, key="clock", help="Used for Live Optics performance sizing")
            st.segmented_control(
                "Live Optics basis", ["95th Percentile", "Peak CPU", "Average CPU"], key="lo_basis",
                disabled=ctx["source"] != "LiveOptics",
                help=f"GHz demand used for performance sizing. '95th' is estimated as {LO_95TH_FALLBACK:.0%} "
                     "of Peak if the column is missing. Live Optics files only.")
            s = st.session_state
            st.caption(f"Host: {s.sockets * s.cores} cores · {s.ram:,} GB RAM  |  "
                       f"NUMA node: {s.cores} cores · {s.ram / s.sockets:,.0f} GB")
        with t2:
            c = st.columns(2)
            c[0].slider("CPU overhead (%)", 0, 50, key="cpu_buffer")
            c[1].slider("RAM overhead (%)", 0, 50, key="ram_buffer")
            c = st.columns(3)
            c[0].number_input("Min cluster size", 1, 32, key="min_hosts")
            c[1].number_input("HA tolerance (hosts)", 0, 4, key="ha_nodes")
        with t3:
            st.toggle("Force a static host count", key="override",
                      help="Use a fixed node count instead of the calculated sizing. The calculated count is still "
                           "shown, and the forced count is checked against the CPU ratio, RAM and HA rules.")
            c = st.columns(3)
            c[0].number_input("Hosts (total, including HA)", 1, 256, key="override_hosts",
                              disabled=not st.session_state.override)
            c[1].toggle("Also use for future (growth) sizing", key="override_future",
                        disabled=not st.session_state.override)
        with t4:
            c = st.columns(3)
            c[0].number_input("Annual growth (%)", 0.0, 100.0, key="growth_pct")
            c[1].number_input("Years", 1, 10, key="years")
            c[2].toggle("Include powered-off VMs", key="include_off")


def host_card(r, which):
    now = which == "now"
    with st.container(border=True):
        eyebrow("Current refresh requirement" if now else f"Future requirement · {r['years']} years")
        hosts = r["hosts_now"] if now else r["hosts_fut"]
        cons = r["constraint"] if now else r["fut_constraint"]
        all_r, fo_r = ((r["ratio_now_all"], r["ratio_now_failover"]) if now
                       else (r["ratio_fut_all"], r["ratio_fut_failover"]))
        sub = (f"Sized for {r['tot_vcpu']:,.0f} vCPU and {r['tot_ram']:,.0f} GB vRAM." if now else
               f"Projected {r['fut_vcpu']:,.0f} vCPU and {r['fut_ram']:,.0f} GB vRAM at "
               f"{r['growth']*100:.0f}% annual growth.")
        big_value(hosts, "hosts", sub)
        tag = host_label(r, which)[1]
        badges = [f":blue-badge[N+{r['ha_nodes']}]", f":gray-badge[{cons} constraint]"]
        if tag:
            badges.append(f":orange-badge[:material/lock: {tag}]")
        st.markdown(" ".join(badges))
        st.markdown(f"vCPU:pCPU {ratio_md(r, all_r)} across all hosts · {ratio_md(r, fo_r)} during HA failover  \n"
                    f":gray[Design limit {r['vcpu_ratio']:.1f}:1 · {hv_label(r)}]")
        if over_hv(r, all_r) or over_hv(r, fo_r):
            st.markdown(f":red-badge[:material/warning: Above {hv_label(r)}]")
        render_checks(r["checks_now"] if now else r["checks_fut"])


def page_recommendation():
    ctx = require_report("Recommendation")
    if not ctx:
        return
    r = ctx["r"]
    page_header("Recommendation", context_subtitle(ctx), ctx)
    scope_bar(ctx)
    render_parameters(ctx)

    section_title("The short answer", "Sizing recommendation")
    c1, c2 = st.columns(2)
    with c1:
        host_card(r, "now")
    with c2:
        host_card(r, "fut")

    section_title("Day 1 design check", f"{r['hosts_now']}-host cluster against the design rules")
    d = r["design_now"]
    cards = [
        ("Target node", f"{r['host_cores']}", "cores", f"{r['ram']:,} GB RAM · {r['cpu_buffer']}% CPU / "
         f"{r['ram_buffer']}% RAM overhead", badge_md(True, "Design input", "", neutral=True)),
        ("CPU at failover", f"{d['cpu_ratio']:.1f}:1", "vCPU / core",
         f"Design limit {r['vcpu_ratio']:.1f}:1 · {hv_label(r)} (failover vCPU:pCPU {fmt_ratio(r['ratio_now_failover'])})",
         badge_md(d["cpu_ok"], "Within limit", "Exceeds limit")
         + (f" :red-badge[Above {r['hypervisor']} {r['hv_kind']}]" if over_hv(r, r["ratio_now_failover"]) else "")),
        ("Memory at failover", f"{d['ram_use']:.0%}", "of usable",
         f"{r['tot_ram']:,.0f} GB vs {d['usable_ram']:,.0f} GB usable", badge_md(d["ram_ok"], "Fits", "Short on RAM")),
        ("Resilience", f"N+{r['ha_nodes']}", f"· {r['hosts_now']} hosts", f"Minimum cluster size {r['min_hosts']}",
         badge_md(d["ha_ok"] and d["min_ok"], "Met", "Not met")),
    ]
    for col, (label, val, unit, sub, badge) in zip(st.columns(4), cards):
        with col, st.container(border=True, height="stretch"):
            eyebrow(label)
            big_value(val, unit, sub)
            st.markdown(badge)

    section_title("Side by side", "Current vs. " + ("proposed" if r["override_now"] else "recommended"))
    day1_col = "Proposed (Day 1)" if r["override_now"] else "Recommended (Day 1)"
    fut_col = f"Future ({r['years']} yrs)"
    view = st.segmented_control("Compare", ["Day 1", "Future", "Both"], key="cmp_view",
                                label_visibility="collapsed") or DEFAULTS["cmp_view"]
    cols = ["Metric", "Current"] + ([day1_col] if view in ("Day 1", "Both") else []) + \
           ([fut_col] if view in ("Future", "Both") else [])
    all_cols = ["Metric", "Current", day1_col, fut_col]
    keep = [all_cols.index(c) for c in cols]
    flagged = comparison_flags(r)
    head = "".join(f"<th>{escape(all_cols[i]) if i else ''}</th>" for i in keep)
    body = "".join(
        "<tr>" + "".join(
            f"<td class='{'hs-over' if (ri, ci) in flagged else ''}{' hs-hl' if ci == 2 else ''}'>{escape(str(row[ci]))}</td>"
            for ci in keep) + "</tr>"
        for ri, row in enumerate(comparison_rows(r)))
    with st.container(border=True):
        st.html(f"<table class='hs-cmp'><tr>{head}</tr>{body}</table>")
        st.caption(f"Ratios in bold red are above the {hv_label(r)}. {r['hv_note']}")

    c1, c2 = st.columns(2)
    with c1, st.container(border=True, height="stretch"):
        st.markdown("#### Target hardware (per node)")
        st.html(kv_html([
            ("CPU", f"{r['sockets']} socket{'' if r['sockets'] == 1 else 's'} x {r['cores']} cores ({r['host_cores']} cores)", False),
            ("RAM", f"{r['ram']:,} GB", False),
            ("Overhead reserved", f"{r['cpu_buffer']}% CPU, {r['ram_buffer']}% RAM", False),
            ("Hypervisor", f"{r['hypervisor']} ({r['hv_ratio']:.1f}:1 {r['hv_kind']})", False),
            ("Design limit", f"{r['vcpu_ratio']:.1f}:1 vCPU:pCPU", r["vcpu_ratio"] > r["hv_ratio"] + EPS),
            ("Min cluster / HA", f"{r['min_hosts']} hosts / N+{r['ha_nodes']}", False),
        ]))
    with c2, st.container(border=True, height="stretch"):
        st.markdown("#### How the number was reached")
        st.markdown("\n".join(f"{i}. {line}" for i, line in enumerate(sizing_logic_lines(r), 1)))

    section_title("Will the largest VMs fit", "Architecture & NUMA")
    numa = numa_messages(r)
    numa_ok = not any(bad for bad, _ in numa)
    c = st.columns(3)
    with c[0], st.container(border=True, height="stretch"):
        eyebrow("Current NUMA node")
        big_value(f"{r['cur_numa_cores']:.0f}", "cores", f"{r['cur_numa_ram']:,.0f} GB per socket today")
    with c[1], st.container(border=True, height="stretch"):
        eyebrow("Target NUMA node")
        big_value(r["tgt_numa_cores"], "cores", f"{r['tgt_numa_ram']:,.0f} GB per socket on the new hosts")
    with c[2], st.container(border=True, height="stretch"):
        eyebrow("Large VM check")
        st.markdown(badge_md(numa_ok, "All VMs fit", "Wide VMs found"))
        for _, text in numa:
            st.caption(text)

    section_title("Per-core licensing", "Licensing impact")
    c = st.columns(3)
    c[0].metric("Current licensed cores", f"{r['cur_lic_cores']:,}", border=True,
                help=f"{r['lic_edition']} · ESXi {r['esxi_version']}")
    c[1].metric("Future licensed cores", f"{r['fut_lic_cores']:,}", border=True,
                help=f"{r['hosts_fut']} hosts x {license_cores_per_host(r['sockets'], r['cores'])} cores")
    c[2].metric("Net change", f"{r['lic_diff']:+,} cores", border=True)
    st.caption(f"Edition: {r['lic_edition']} · ESXi: {r['esxi_version']} · "
               f"{LIC_MIN_CORES_PER_SOCKET}-core per CPU minimum applied")

    ins = perf_insight(r)
    if ins:
        good, title, msg = ins
        section_title("Measured, not allocated", "Performance sizing (Live Optics)")
        c = st.columns(3)
        c[0].metric("Allocated vCPU", f"{r['tot_vcpu']:,.0f}", border=True)
        c[1].metric(f"Consumed GHz ({r['lo_basis']})", f"{r['perf_ghz']:,.1f}", border=True)
        c[2].metric("Performance-based hosts", r["perf_hosts_rec"], border=True)
        st.markdown(badge_md(good, title, title) + f" {msg}")
        if r["perf_estimated"]:
            st.caption(f"95th percentile column not found; estimated as {LO_95TH_FALLBACK:.0%} of peak.")


def _store_logo():
    f = st.session_state.get("logo_uploader")
    if f is not None:
        st.session_state.cust_logo = {"bytes": f.getvalue(), "type": f.type or "image/png", "name": f.name}


def page_report():
    ctx = require_report("Report")
    if not ctx:
        return
    page_header("Report", "AHEAD-branded HTML report. Self-contained: open it in any browser, email it, "
                          "or print to PDF.", ctx)
    c1, c2 = st.columns([2, 1])
    with c1, st.container(border=True, height="stretch"):
        st.markdown("#### Report settings")
        st.text_input("Customer name", key="customer")
        st.file_uploader("Customer logo (optional)", type=["png", "jpg", "jpeg", "svg"], key="logo_uploader",
                         on_change=_store_logo, help="Shown in the report's 'Prepared for' chip.")
    with c2, st.container(border=True, height="stretch"):
        st.markdown("#### Customer logo")
        logo = st.session_state.cust_logo
        if logo:
            st.image(logo["bytes"], width=180)
            st.button("Remove logo", icon=":material/delete:",
                      on_click=lambda: st.session_state.update(cust_logo=None))
        else:
            st.caption("No logo uploaded. The customer name is shown instead.")
    section_title("Preview", "What the customer will see")
    if hasattr(st, "iframe"):
        st.iframe(ctx["html"], height=1400)
    else:
        components.html(ctx["html"], height=1400, scrolling=True)


def page_raw():
    ctx = require_report("Raw Data")
    if not ctx:
        return
    page_header("Raw Data", context_subtitle(ctx), ctx)
    c1, c2, c3 = st.columns([1, 1, 2], vertical_alignment="bottom")
    view = c1.segmented_control("Dataset", ["VMs", "Hosts"], key="raw_view") or DEFAULTS["raw_view"]
    cols_mode = c2.segmented_control("Columns", ["Key columns", "All columns"], key="raw_cols") or DEFAULTS["raw_cols"]
    query = c3.text_input("Search", key="raw_search", icon=":material/search:", placeholder="Filter rows...")

    df = ctx["inv"]["raw_vms"] if view == "VMs" else ctx["inv"]["raw_hosts"]
    if cols_mode == "Key columns":
        keep = [c for c in KEY_COLUMNS[ctx["source"]][view] if c in df.columns]
        if keep:
            df = df[keep]
    if query:
        mask = df.astype(str).apply(lambda s: s.str.contains(query, case=False, na=False, regex=False)).any(axis=1)
        df = df[mask]
    name_col = first_col(df, ["VM", "VM Name", "Host", "Host Name"])
    st.caption(f"{len(df):,} {(view[:-1] if len(df) == 1 else view)} in scope" + (f" matching '{query}'" if query else "")
               + (" · saved project: only the columns used for sizing are kept"
                  if st.session_state.reports[ctx["key"]].get("sheets") is not None else ""))
    st.dataframe(df, hide_index=True, width="stretch", height=560,
                 column_config={name_col: st.column_config.TextColumn(pinned=True)} if name_col else None)


# ---------- entry point ----------
def main():
    global CTX
    st.set_page_config(page_title=f"{PRODUCT_NAME} | AHEAD", layout="wide", page_icon=":material/dns:")
    version = tuple(int(x) for x in st.__version__.split(".")[:2] if x.isdigit())
    if version < MIN_STREAMLIT:
        st.error(f"{PRODUCT_NAME} needs Streamlit {'.'.join(map(str, MIN_STREAMLIT))} or newer; this computer has "
                 f"{st.__version__}. Start the tool with Run_on_mac.command or Run_on_windows.bat, which update it "
                 f"automatically.")
        st.stop()
    if ensure_theme_config():
        st.toast("AHEAD theme installed. Restart the app to apply it.", icon=":material/palette:")
    init_state()
    st.html(APP_CSS)

    PAGES.update({
        "home": st.Page(page_home, title="Home", icon=":material/home:", url_path="home", default=True),
        "env": st.Page(page_environment, title="Current Environment", icon=":material/dns:", url_path="environment"),
        "rec": st.Page(page_recommendation, title="Recommendation", icon=":material/insights:",
                       url_path="recommendation"),
        "report": st.Page(page_report, title="Report", icon=":material/description:", url_path="report"),
        "raw": st.Page(page_raw, title="Raw Data", icon=":material/table_view:", url_path="raw-data"),
    })
    nav = st.navigation(list(PAGES.values()), position="hidden")

    goto = st.session_state.pop("_goto", None)
    if goto:
        st.switch_page(PAGES[goto])

    CTX = build_context()
    render_sidebar()
    nav.run()


if __name__ == "__main__":
    if "--setup" in sys.argv:
        ensure_theme_config()
        print("AHEAD theme config is in place.")
    else:
        main()
