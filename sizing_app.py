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
import io
import math
import os
from datetime import datetime
from html import escape

import pandas as pd
import streamlit as st
import streamlit.components.v1 as components

# =============================================================================
# 1. CONSTANTS
# =============================================================================
APP_TITLE = "Virtualization Sizing Calculator"
APP_VERSION = "2.4"
APP_DIR = os.path.dirname(os.path.abspath(__file__))
DEFAULT_LOGO_PATH = os.path.join(APP_DIR, "ahead_logo.png")          # navy, light backgrounds
WHITE_LOGO_PATH = os.path.join(APP_DIR, "ahead_logo_white.png")      # reversed, dark backgrounds

ALL_CLUSTERS = "All Clusters"
LIC_MIN_CORES_PER_SOCKET = 16   # VMware per-core licensing minimum per CPU
PERF_TARGET_UTIL = 0.80         # Performance sizing keeps hosts at <= 80% CPU
LO_95TH_FALLBACK = 0.95         # 95th estimate = 95% of peak when column missing
MIB_PER_TB = 1024 * 1024
EPS = 1e-9                      # guards ceil() against float noise (e.g. 3.0000000001)

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
        ("Hosts", f"{r['cur_host_count']:,}"),
        ("Physical Cores", f"{r['cur_cores']:,.0f}"),
        ("Host RAM", f"{r['cur_total_ram_gb']:,.0f} GB"),
        ("vCPU:pCPU", fmt_ratio(r["cur_ratio"])),
        ("VMs", f"{r['tot_vms']:,}"),
        ("vCPU", f"{r['tot_vcpu']:,.0f}"),
        ("vRAM", f"{r['tot_ram']:,.0f} GB"),
        ("vRAM / Host RAM", _pct(r["cur_mem_alloc"])),
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
        ("vCPU:pCPU (all hosts)", fmt_ratio(r["cur_ratio"]), r["cur_ratio"] > r["vcpu_ratio"]),
        ("vCPU:pCPU (1 host failed)", fmt_ratio(r["cur_ratio_n1"]),
         r["cur_ratio_n1"] is not None and r["cur_ratio_n1"] > r["vcpu_ratio"]),
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
            val = f"<strong>&#9888; {e(v)}</strong>" if warn else e(v)
            cells.append(f"<tr><td>{e(k)}</td><td>{val}</td></tr>")
        return "<table class='kv'>" + "".join(cells) + "</table>"

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
        (fmt_ratio(r["ratio_now_failover"]), f"vCPU:pCPU during HA failover · design limit {r['vcpu_ratio']:.1f}:1"),
    ]
    stats_html = "".join(f"<div class='stat'><div class='num'>{e(n)}</div><div class='cap'>{e(c)}</div></div>"
                         for n, c in stats)
    hero_lede = (f"Host sizing for {r['scope_label']} from the {r['source_type']} export. Today "
                 f"{plural(r['cur_host_count'], 'host')} with {r['cur_cores']:,.0f} cores and "
                 f"{r['cur_total_ram_gb']:,.0f} GB RAM run {plural(r['tot_vms'], 'VM')}. "
                 f"The target is {plural(r['hosts_now'], 'node')} of {r['sockets']} x {r['cores']}-core, "
                 f"{r['ram']:,} GB hosts, growing to {r['hosts_fut']} over {r['years']} years.")

    # ---------- Section 1: current environment ----------
    kpi_html = "".join(
        f"<div class='card kpi {'navy' if i < 4 else ''}'><div class='label'>{e(k)}</div><div class='big'>{e(v)}</div></div>"
        for i, (k, v) in enumerate(headline))
    groups_html = "".join(
        f"<div class='card {'green' if i % 2 else ''}'><div class='label'>{e(title)}</div>{kv_table(rows)}</div>"
        for i, (title, rows) in enumerate(groups))
    cur_lede = (f"{plural(r['cur_host_count'], 'host')} with {r['cur_cores']:,.0f} physical cores and "
                f"{r['cur_total_ram_gb']:,.0f} GB RAM run {plural(r['tot_vms'], 'VM')} "
                f"({r['tot_vcpu']:,.0f} vCPU, {r['tot_ram']:,.0f} GB vRAM) at {fmt_ratio(r['cur_ratio'])} vCPU:pCPU"
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
                  + f" vCPU:pCPU {fmt_ratio(all_r)} across all hosts, {fmt_ratio(fo_r)} during HA failover.")
        return (f"<div class='card {'green lead' if now else ''}'><div class='label'>{e(title)}</div>"
                f"<div class='big'>{e(plural(hosts, 'host'))}</div>"
                + (f"<span class='tag'>{e(tag)}</span>" if tag else "")
                + f"<p>{detail}</p>{checks_list(r['checks_now' if now else 'checks_fut'])}</div>")

    d = r["design_now"]
    panel_cards = [
        ("Target node", f"{r['host_cores']} cores · {r['ram']:,} GB",
         f"{r['sockets']} sockets x {r['cores']} cores. {r['cpu_buffer']}% CPU and {r['ram_buffer']}% RAM reserved for overhead.",
         pill(True, "Design input", "info")),
        ("CPU", f"{d['cpu_ratio']:.1f}:1 at failover",
         f"vCPU per usable core with {plural(r['ha_nodes'], 'host')} down. Design limit {r['vcpu_ratio']:.1f}:1.",
         pill(d["cpu_ok"], "Within limit" if d["cpu_ok"] else "Exceeds limit")),
        ("Memory", f"{d['ram_use']:.0%} of usable RAM",
         f"{r['tot_ram']:,.0f} GB vRAM against {d['usable_ram']:,.0f} GB usable during HA failover.",
         pill(d["ram_ok"], "Fits" if d["ram_ok"] else "Short on RAM")),
        ("Resilience", f"N+{r['ha_nodes']} · {plural(r['hosts_now'], 'host')}",
         f"Minimum cluster size {r['min_hosts']}. {plural(max(d['active'], 0), 'host')} carry the workload after a failure.",
         pill(d["ha_ok"] and d["min_ok"], "Met" if d["ha_ok"] and d["min_ok"] else "Not met")),
    ]
    panel_html = "".join(
        f"<div class='pcard'><div class='eyebrow'>{e(a)}</div><h4>{e(b)}</h4><p>{e(c)}</p>{p}</div>"
        for a, b, c, p in panel_cards)

    compare_html = (f"<table class='compare'><tr><th></th><th>Current</th><th>{e(day1_col)}</th>"
                    f"<th>Future ({r['years']} yrs)</th></tr>"
                    + "".join(f"<tr><td>{e(a)}</td><td>{e(b)}</td><td class='hl'>{e(c)}</td><td>{e(dd)}</td></tr>"
                              for a, b, c, dd in comparison_rows(r)) + "</table>")
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
  <div class="eyebrow kicker">Virtualization sizing · {e(r['source_type'])} analysis · {now:%b %d, %Y}</div>
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
def logo_data_uri(uploaded_logo):
    """Embed the customer's logo in the report so it works offline / when emailed."""
    if uploaded_logo is None:
        return None
    mime = uploaded_logo.type or "image/png"
    return f"data:{mime};base64,{base64.b64encode(uploaded_logo.getvalue()).decode()}"


def sidebar_params():
    sb = st.sidebar
    if os.path.exists(DEFAULT_LOGO_PATH):
        sb.image(DEFAULT_LOGO_PATH, width=140)
    sb.title("⚙️ Sizing Parameters")
    sb.subheader("1. Target Hardware")
    p = {
        "sockets": sb.number_input("Sockets/Host", 1, 4, 2),
        "cores": sb.number_input("Cores/Socket", 4, 128, 24),
        "ram": sb.number_input("RAM/Host (GB)", 64, 8192, 1024),
        "clock": sb.number_input("CPU Speed (GHz)", 1.0, 5.0, 2.5, help="Used for Live Optics performance sizing"),
    }
    p["lo_basis"] = sb.radio(
        "Live Optics Basis", ["95th Percentile", "Peak CPU", "Average CPU"], index=0,
        help=f"GHz demand used for performance sizing. '95th' is estimated as {LO_95TH_FALLBACK:.0%} of Peak if the column is missing.")
    sb.info(f"**Host Spec:**\n{p['sockets'] * p['cores']} Cores | {p['ram']} GB RAM\n\n"
            f"**NUMA Node:**\n{p['cores']} Cores | {p['ram'] / p['sockets']:.0f} GB")

    sb.subheader("2. Constraints")
    p["vcpu_ratio"] = sb.slider("Max vCPU:pCPU (Design Limit)", 1.0, 10.0, 5.0, 0.5)
    p["cpu_buffer"] = sb.slider("CPU Overhead (%)", 0, 50, 10)
    p["ram_buffer"] = sb.slider("RAM Overhead (%)", 0, 50, 10)
    p["min_hosts"] = sb.number_input("Min Cluster Size", 1, 32, 2)
    p["ha_nodes"] = sb.number_input("HA Tolerance", 0, 4, 1)

    sb.subheader("3. Host Count Override")
    p["override"] = sb.checkbox(
        "Force a static host count",
        help="Use a fixed node count instead of the calculated sizing. The calculated count is still shown, "
             "and the forced count is checked against the CPU ratio, RAM and HA rules.")
    p["override_hosts"] = sb.number_input("Hosts (total, including HA)", 1, 256, 4, disabled=not p["override"])
    p["override_future"] = sb.checkbox("Also use for future (growth) sizing", value=True,
                                       disabled=not p["override"])

    sb.subheader("4. Scope & Growth")
    p["include_off"] = sb.checkbox("Include Powered Off VMs?", value=True)
    p["growth"] = sb.number_input("Annual Growth (%)", 0.0, 100.0, 10.0) / 100
    p["years"] = sb.number_input("Years", 1, 10, 3)

    sb.divider()
    sb.subheader("📥 Report Settings")
    p["customer"] = sb.text_input("Customer Name", "My Customer")
    logo_file = sb.file_uploader("Customer logo (optional)", type=["png", "jpg", "jpeg", "svg"],
                                 help="Shown in the report's 'Prepared for' chip. AHEAD branding is applied automatically.")
    logo_uri = logo_data_uri(logo_file)
    if logo_file:
        sb.image(logo_file, width=150)
    return p, logo_uri


def _kv_markdown(rows):
    lines = ["| | |", "|---|---|"]
    for k, v, warn in rows:
        v = str(v).replace("|", "\\|")
        lines.append(f"| {k} | {'⚠️ ' if warn else ''}**{v}** |")
    return "\n".join(lines)


def render_checks(checks):
    for bad, text in checks:
        (st.error if bad else st.caption)(("⚠️ " if bad else "✅ ") + text)


def render_summary(r):
    # ---- 1. Current environment (first) ----
    headline, groups = current_env(r)
    st.subheader(f"1. Current Environment ({r['scope_label']})")
    for row in (headline[:4], headline[4:]):
        for col, (label, value) in zip(st.columns(4), row):
            col.metric(label, value)
    for pair in (groups[:2], groups[2:]):
        for col, (title, rows) in zip(st.columns(2), pair):
            with col, st.container(border=True):
                st.markdown(f"#### {title}")
                st.markdown(_kv_markdown(rows))

    # ---- 2. Sizing recommendation ----
    st.subheader("2. Sizing Recommendation")
    c1, c2 = st.columns(2)
    with c1:
        st.info("### Current Refresh Requirement", icon="📅")
        st.markdown(f"### **{r['hosts_now']} Nodes** (N+{r['ha_nodes']} HA)")
        if r["override_now"]:
            st.caption(f"🔒 {host_label(r, 'now')[1]}")
        st.write(f"**vCPU:pCPU:** {fmt_ratio(r['ratio_now_all'])} all hosts | "
                 f"{fmt_ratio(r['ratio_now_failover'])} during HA failover (design limit {r['vcpu_ratio']:.1f}:1)")
        if r["constraint"] == "CPU":
            st.success("✅ **Constraint: CPU Bound**")
        elif r["constraint"] == "RAM":
            st.warning("⚠️ **Constraint: Memory Bound**")
        else:
            st.info("**Constraint: CPU and RAM balanced**")
        render_checks(r["checks_now"])
    with c2:
        st.success("### Future Requirement with Growth", icon="🚀")
        st.markdown(f"### **{r['hosts_fut']} Nodes** (+{r['growth']*100:.0f}% / {r['years']} yrs)")
        if r["override_fut"]:
            st.caption(f"🔒 {host_label(r, 'fut')[1]}")
        st.write(f"**vCPU:pCPU:** {fmt_ratio(r['ratio_fut_all'])} all hosts | "
                 f"{fmt_ratio(r['ratio_fut_failover'])} during HA failover")
        st.write(f"**Constraint:** {r['fut_constraint']}")
        render_checks(r["checks_fut"])

    with st.container(border=True):
        day1_col = "Proposed (Day 1)" if r["override_now"] else "Recommended (Day 1)"
        st.markdown(f"#### Current vs. {day1_col.split(' (')[0]}")
        st.table(pd.DataFrame(comparison_rows(r), columns=["Metric", "Current", day1_col,
                                                           f"Future ({r['years']} yrs)"]).set_index("Metric"))

    t1, t2 = st.columns(2)
    with t1, st.container(border=True):
        st.markdown("#### Target Hardware (per node)")
        st.markdown(_kv_markdown([
            ("CPU", f"{r['sockets']} sockets x {r['cores']} cores ({r['host_cores']} cores)", False),
            ("RAM", f"{r['ram']:,} GB", False),
            ("Overhead reserved", f"{r['cpu_buffer']}% CPU, {r['ram_buffer']}% RAM", False),
            ("Design limit", f"{r['vcpu_ratio']:.1f}:1 vCPU:pCPU", False),
            ("Min cluster / HA", f"{r['min_hosts']} hosts / N+{r['ha_nodes']}", False),
        ]))
    with t2, st.container(border=True):
        st.markdown("#### 📊 Sizing Logic")
        for i, line in enumerate(sizing_logic_lines(r), 1):
            st.write(f"**{i}.** {line}")

    # ---- 3. NUMA ----
    st.subheader("3. Architecture & NUMA")
    st.write(f"**Current NUMA:** {r['cur_numa_cores']:.0f} Cores | {r['cur_numa_ram']:,.0f} GB  →  "
             f"**Target NUMA:** {r['tgt_numa_cores']} Cores | {r['tgt_numa_ram']:,.0f} GB")
    for bad, text in numa_messages(r):
        (st.warning if bad else st.success)(("⚠️ " if bad else "✅ ") + text)

    # ---- 4. Licensing ----
    st.subheader("4. Licensing Impact")
    l1, l2, l3 = st.columns(3)
    l1.metric("Current Licensed Cores", f"{r['cur_lic_cores']:,}")
    l2.metric("Future Licensed Cores", f"{r['fut_lic_cores']:,}")
    l3.metric("Net Change", f"{r['lic_diff']:+,} cores")
    st.caption(f"Edition: {r['lic_edition']} | ESXi: {r['esxi_version']} | "
               f"{LIC_MIN_CORES_PER_SOCKET}-core per CPU minimum applied")

    # ---- 5. Performance ----
    ins = perf_insight(r)
    if ins:
        good, title, msg = ins
        st.subheader("5. Performance Sizing (Live Optics)")
        p1, p2, p3 = st.columns(3)
        p1.metric("Allocated vCPU", f"{r['tot_vcpu']:,.0f}")
        p2.metric(f"Consumed GHz ({r['lo_basis']})", f"{r['perf_ghz']:,.1f}")
        p3.metric("Performance-Based Hosts", r["perf_hosts_rec"])
        (st.success if good else st.error)(f"**{title}:** {msg}")
        if r["perf_estimated"]:
            st.caption(f"95th percentile column not found; estimated as {LO_95TH_FALLBACK:.0%} of peak.")


def main():
    st.set_page_config(page_title=APP_TITLE, layout="wide", page_icon="📊")
    params, logo_uri = sidebar_params()

    st.title(f"📊 {APP_TITLE}")
    st.markdown(f"Automated hardware sizing analysis for **{params['customer']}**.")

    files = st.file_uploader("Upload RVTools or Live Optics Excel file(s)", type=["xlsx"],
                             accept_multiple_files=True, label_visibility="collapsed")
    if not files:
        st.info("👆 Upload one or more **RVTools (.xlsx)** or **Live Optics (.xlsx)** files to begin.")
        st.stop()

    if len(files) > 1:
        names = [f.name for f in files]
        active = files[names.index(st.selectbox("Active report", names))]
    else:
        active = files[0]

    try:
        source_type, sheets = load_workbook(active.getvalue())
    except Exception as exc:
        st.error(f"Could not read '{active.name}'. Is it a valid .xlsx export?")
        with st.expander("Error details"):
            st.exception(exc)
        st.stop()

    if source_type is None:
        st.error("Unknown file. Expected an RVTools export (vInfo tab) or Live Optics export (VMs + ESX Hosts tabs).")
        st.stop()
    st.success(f"📂 **{source_type}** detected: {active.name}")

    clusters = cluster_list(source_type, sheets)
    chosen = st.multiselect(
        "Select Cluster Scope (choose 2+ to model a consolidation):",
        options=[ALL_CLUSTERS] + clusters, default=[ALL_CLUSTERS],
        help="'All Clusters' for the full environment, one cluster for a standalone sizing, "
             "or several clusters to size a consolidated target.") if clusters else []
    scope = None if (not chosen or ALL_CLUSTERS in chosen) else chosen
    scope_label = ALL_CLUSTERS if scope is None else " + ".join(scope)
    if scope and len(scope) > 1:
        st.info(f"🧩 **Consolidation Scope:** Combining {len(scope)} clusters: {scope_label}")

    try:
        if source_type == "RVTools":
            inv = parse_rvtools(sheets, scope, params["include_off"])
        else:
            inv = parse_live_optics(sheets, scope, params["include_off"], params["lo_basis"])
        r = compute_sizing(inv, params)
        r.update(scope_label=scope_label, file_name=active.name, source_type=source_type)
        report_html = generate_html_report(r, logo_uri)
    except KeyError as exc:
        st.error(f"A required column or tab is missing from the file: {exc}")
        with st.expander("Error details"):
            st.exception(exc)
        st.stop()
    except Exception as exc:
        st.error(f"Error processing file: {exc}")
        with st.expander("Error details"):
            st.exception(exc)
        st.stop()

    base = os.path.splitext(active.name)[0]
    safe_scope = "".join(ch if ch.isalnum() or ch in "-_" else "_" for ch in scope_label.replace(" + ", "-"))
    st.sidebar.download_button("Download Report", report_html,
                               file_name=f"{base}_{safe_scope}_Sizing.html", mime="text/html")

    tab1, tab2, tab3 = st.tabs(["📋 Executive Report", "🖨️ Report Preview", "🔍 Raw Data"])
    with tab1:
        render_summary(r)
    with tab2:
        if hasattr(st, "iframe"):           # Streamlit >= 1.52
            st.iframe(report_html, height=1400)
        else:
            components.html(report_html, height=1400, scrolling=True)
    with tab3:
        st.write(f"**VMs in scope** ({len(inv['raw_vms']):,} rows)")
        st.dataframe(inv["raw_vms"])
        st.write(f"**Hosts in scope** ({len(inv['raw_hosts']):,} rows)")
        st.dataframe(inv["raw_hosts"])


if __name__ == "__main__":
    main()
