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
import copy
import gzip
import hashlib
import io
import json
import math
import os
import re
import sys
import uuid
from datetime import datetime
from html import escape

import pandas as pd
import streamlit as st
import streamlit.components.v1 as components

# =============================================================================
# 1. CONSTANTS
# =============================================================================
APP_TITLE = "Virtualization Sizing Calculator"
APP_VERSION = "3.13"
APP_DIR = os.path.dirname(os.path.abspath(__file__))
DEFAULT_LOGO_PATH = os.path.join(APP_DIR, "ahead_logo.png")          # navy, light backgrounds
WHITE_LOGO_PATH = os.path.join(APP_DIR, "ahead_logo_white.png")      # reversed, dark backgrounds
PRODUCT_LOGO_PATH = os.path.join(APP_DIR, "host_sizer_logo_white.png")  # Host Sizer lockup (dark sidebar)
PRODUCT_ICON_PATH = os.path.join(APP_DIR, "host_sizer_icon.png")        # browser-tab icon

ALL_CLUSTERS = "All Clusters"
LIC_MIN_CORES_PER_SOCKET = 16   # VMware per-core licensing minimum per CPU
PERF_TARGET_UTIL = 0.80         # Performance sizing keeps hosts at <= 80% CPU
LO_95TH_FALLBACK = 0.95         # 95th estimate = 95% of peak when column missing
MIB_PER_TB = 1e6 / 1.048576        # binary MiB per decimal TB (reports use decimal TB)
EPS = 1e-9                      # guards ceil() against float noise (e.g. 3.0000000001)

# Hypervisor vCPU:pCPU guidance: (ratio, "average"/"maximum", source note)
HYPERVISORS = {
    "VMware": (4.5, "average", "Average for the workloads Intel has captured."),
    "Nutanix": (3.5, "average", "Average vCPU:pCPU ratio."),
    "Azure Local": (2.5, "maximum", "Maximum Intel recommends; Dell product management agrees."),
    "Hyper-V": (2.0, "average", "Average vCPU:pCPU ratio."),
}
DEFAULT_HYPERVISOR = "VMware"

RVTOOLS_SHEETS = ["vInfo", "vHost", "vDatastore", "vPartition", "vDisk", "vLicense", "vSnapshot", "vMemory", "vMetaData", "vMultiPath"]
LIVEOPTICS_SHEETS = ["VMs", "ESX Hosts", "Host Devices", "ESX Performance", "ESX Licenses", "Host Disks"]

HOST_NAME_COLS = ["Host Name", "Host", "ESX Host", "Hostname", "Server Name"]
LUN_ID_COLS = ["Canonical Name", "Device ID", "Device UUID", "UUID", "NAA ID", "Device Name"]
RDM_DISK_COLS = ["VM", "Raw", "Disk", "Label", "Raw LUN ID", "Raw Comp. Mode", "Disk Mode", "Sharing mode",
                 "Shared Bus", "Controller", "Host", "Cluster"]
LUN_COLS = ["Host", "Cluster", "Datastore", "Disk", "Display name", "Vendor", "Model", "UUID"]

# ---------- storage consumption (3.13) ----------
# Exports report binary units (RVTools "MiB"; its older "MB" columns are binary too; Live Optics MiB/GiB).
# Everything is converted to decimal MB and TB (1 TB = 1,000,000 MB), the units storage arrays are sold in.
MIB_MB, GIB_MB, TIB_MB = 1.048576, 1073.741824, 1099511.627776
UNIT_MB = (("MiB", MIB_MB), ("MB", MIB_MB), ("GiB", GIB_MB), ("GB", GIB_MB), ("TiB", TIB_MB), ("TB", TIB_MB))
SNAP_SHEET = "Snapshots"                 # Live Optics snapshot tab, whatever it is called in the export
SNAP_VM_COLS = ["VM", "VM Name", "Virtual Machine", "Name"]
SNAP_DATE_COLS = ["Date / time", "Date/time", "Date / Time", "Created", "Create Time", "Creation Date",
                  "Creation Time", "Date"]
SNAP_COLS = SNAP_VM_COLS + SNAP_DATE_COLS + [f"Size {u}{s}" for u, _ in UNIT_MB for s in (" (total)", " (vmsn)", "")] + \
    [f"Size ({u})" for u, _ in UNIT_MB] + [f"Snapshot Size ({u})" for u, _ in UNIT_MB] + ["Cluster", "Powerstate"]
EMPTY_STO = {"vm_alloc": 0.0, "vm_inuse": 0.0, "guest": 0.0, "swap": 0.0, "tmpl_cnt": 0, "tmpl_alloc": 0.0,
             "tmpl_inuse": 0.0, "rdm_cnt": 0, "rdm": 0.0, "snap_mb": 0.0, "snap_cnt": 0, "snap_vms": 0,
             "ds_cap": 0.0, "ds_used": 0.0, "ds_alloc": 0.0, "other": None, "has_snap": False, "has_swap": False,
             "snap_dated": False, "snap_ref": None, "notes": [], "vsan_src": False, "vsan_vms": 0,
             "vsan_vm_alloc": 0.0, "vsan_vm_inuse": 0.0, "vsan_tmpl_alloc": 0.0, "vsan_tmpl_inuse": 0.0,
             "vsan_ds": 0, "vsan_cap": 0.0, "vsan_used": 0.0, "vsan_raw": 0.0, "vsan_vm_used_logical": 0.0,
             "vsan_layout": "", "vsan_kind": "", "vsan_vxrail": False, "vsan_copies_in_vm": True,
             "rdms": None, "host_luns": None}




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
    """RVTools: find '{base} <unit>' and convert to decimal TB."""
    return mb_sum(df, base) / 1e6


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
    if source == "LiveOptics":                 # snapshot tab name varies between Live Optics versions
        snap = next((n for n in xls.sheet_names if "snapshot" in n.lower()), None)
        if snap:
            df = xls.parse(snap)
            df.columns = df.columns.astype(str).str.strip()
            sheets[SNAP_SHEET] = df
    return source, fill_standalone(source, sheets)


STANDALONE = "Standalone hosts"
STANDALONE_SHEETS = {"RVTools": ["vInfo", "vHost", "vMultiPath"],
                     "LiveOptics": ["VMs", "ESX Hosts", "ESX Performance", "Host Devices"]}


def fill_standalone(source, sheets):
    """VMs and hosts with no cluster (stand-alone ESXi hosts) become one 'Standalone hosts' group, so they
    can be selected and sized like a cluster instead of being silently left out."""
    for name in STANDALONE_SHEETS.get(source, []):
        df = sheets.get(name)
        if df is None or df.empty:
            continue
        if "Cluster" not in df.columns:
            if name == "Host Devices":       # filtered by host name instead
                continue
            df = df.assign(Cluster=STANDALONE)
        else:
            blank = df["Cluster"].isna() | df["Cluster"].astype(str).str.strip().isin(["", "nan", "None"])
            if not blank.any():
                continue
            df = df.copy()
            df["Cluster"] = df["Cluster"].astype(object)
            df.loc[blank, "Cluster"] = STANDALONE
        sheets[name] = df
    return sheets


def cluster_list(source, sheets):
    df = sheets["vInfo"] if source == "RVTools" else sheets["VMs"]
    if "Cluster" not in df.columns:
        return []
    return sorted(df["Cluster"].dropna().astype(str).unique())


# ---------- 3.11: power scope, headroom, method section ----------
POWER_ALL, POWER_ON, POWER_OFF = "All VMs", "Powered on only", "Powered off only"
POWER_MODES = [POWER_ALL, POWER_ON, POWER_OFF]


def power_mode(values):
    """Normalise the VM power-state choice. Older settings stored include_off (True/False)."""
    v = (values or {}).get("power_scope")
    if v in POWER_MODES:
        return v
    return POWER_ALL if (values or {}).get("include_off", True) else POWER_ON


def filter_power(df, col, mode):
    """Keep all VMs, only powered-on, or only powered-off (suspended counts as off)."""
    if mode == POWER_ALL or col not in df.columns:
        return df
    on = df[col].astype(str).str.contains("poweredOn", case=False, na=False)
    return df[on] if mode == POWER_ON else df[~on]


def vm_headroom(r, hosts, vcpu, ram, ghz=0.0):
    """How many more average-sized VMs fit on `hosts` within the design rules (HA hosts failed,
    vCPU:pCPU design limit, usable RAM after overhead) before another host is needed."""
    if not r["tot_vms"] or (not r["tot_vcpu"] and not r["tot_ram"]):
        return None
    avg_c, avg_m = r["tot_vcpu"] / r["tot_vms"], r["tot_ram"] / r["tot_vms"]
    avg_g = r["perf_ghz"] / r["tot_vms"] if r.get("has_perf") else 0.0
    active = max(hosts - r["ha_nodes"], 0)
    rooms = {"CPU": (active * r["eff_cores"] * r["vcpu_ratio"] - vcpu) / avg_c if avg_c else math.inf,
             "RAM": (active * r["eff_ram"] - ram) / avg_m if avg_m else math.inf}
    if avg_g:
        rooms["GHz"] = (active * r["eff_ghz"] - ghz) / avg_g
    limit = min(rooms, key=rooms.get)
    return {"vms": math.floor(rooms[limit] + EPS), "limit": limit}


def add_headroom(r):
    if not r["tot_vms"]:
        r.update(avg_vm_vcpu=0.0, avg_vm_ram=0.0, avg_vm_ghz=0.0, headroom_now=None, headroom_fut=None, vms_per_host=None)
        return
    avg_c, avg_m = r["tot_vcpu"] / r["tot_vms"], r["tot_ram"] / r["tot_vms"]
    avg_g = r["perf_ghz"] / r["tot_vms"] if r.get("has_perf") else 0.0
    per_c = r["eff_cores"] * r["vcpu_ratio"] / avg_c if avg_c else math.inf
    per_m = r["eff_ram"] / avg_m if avg_m else math.inf
    per_g = r["eff_ghz"] / avg_g if avg_g else math.inf
    per = min(per_c, per_m, per_g)
    r.update(avg_vm_vcpu=avg_c, avg_vm_ram=avg_m, avg_vm_ghz=avg_g,
             headroom_now=vm_headroom(r, r["hosts_now"], r["tot_vcpu"], r["tot_ram"], r["perf_ghz"]),
             headroom_fut=vm_headroom(r, r["hosts_fut"], r["fut_vcpu"], r["fut_ram"], r.get("fut_ghz", 0.0)),
             vms_per_host=math.floor(per + EPS) if per != math.inf else None)


def headroom_text(h):
    if h is None:
        return "n/a"
    if h["vms"] < 0:
        return f"Over by {-h['vms']:,} VMs"
    return f"{h['vms']:,} more VMs"


METHOD_STEPS = [
    ("Data", "Inventory from RVTools (configured vCPU and memory) and/or Live Optics (which also measures CPU demand). "
             "It is a point-in-time snapshot."),
    ("Scope", "Templates and placeholder VMs are excluded. Either all VMs, only powered-on VMs, or only powered-off VMs "
              "(for DR) are sized. Clusters marked out of scope are left out."),
    ("Usable host", "Each new host's cores and memory, minus the CPU and memory overhead reserved for the hypervisor."),
    ("Host count", "Hosts needed for CPU (vCPU ÷ the vCPU:pCPU design limit) and for memory; the larger number wins. "
                   "HA spare hosts are added, and the cluster never goes below the minimum size. The future figure repeats "
                   "this with compound annual growth."),
    ("Checks", "The design is re-tested with the HA hosts failed. vCPU:pCPU ratios above the hypervisor's guidance are "
               "flagged in red."),
    ("Room to grow", "How many more VMs of the current average size fit within the same rules before another host is needed."),
    ("Also shown", "Live Optics performance sizing (80% CPU target), NUMA fit for large VMs, and per-core licensing "
                   "(16-core minimum per CPU)."),
]
METHOD_NOTE = ("This is a planning estimate based on the data provided. Storage is reported as consumption and "
               "effective capacity only; vSAN, array protection and data reduction, GPU and application-specific "
               "requirements are not included and should be confirmed during design.")


def method_section_html(rs, section_class=""):
    """Short, general explanation of the sizing method plus the settings used for each cluster."""
    e = lambda v: escape(str(v))  # noqa: E731
    steps = "".join(f"<li><strong>{e(t)}.</strong> {e(d)}</li>" for t, d in METHOD_STEPS)
    rows = "".join(
        f"<tr><td>{e(r['scope_label'])}</td><td>{e(r['hypervisor'])} ({r['hv_ratio']:.1f}:1)</td>"
        f"<td>{r['vcpu_ratio']:.1f}:1</td><td>{r['cpu_buffer']}% / {r['ram_buffer']}%</td>"
        f"<td>N+{r['ha_nodes']} · min {r['min_hosts']}</td><td>{r['growth']*100:.0f}% × {r['years']} yrs</td>"
        f"<td>{e(r.get('power_scope', POWER_ALL))}</td>"
        f"<td>{e(str(r['override_hosts']) + ' hosts' if r['override_now'] else '—')}</td></tr>" for r in rs)
    return f"""
<section class="{section_class}" id="method"><div class="wrap">
  <div class="eyebrow">How this sizing was done</div>
  <h2>Sizing Method</h2>
  <div class="card navy" style="margin-top:28px"><ol class="logic" style="padding-left:18px">{steps}</ol></div>
  <h3>Settings used</h3>
  <div class="table-card scroll" style="margin-top:0"><table class="compare">
    <tr><th>Cluster</th><th>Hypervisor</th><th>Design limit</th><th>Overhead CPU / RAM</th><th>HA</th>
        <th>Growth</th><th>VMs sized</th><th>Override</th></tr>{rows}</table></div>
  <div class="note">{e(METHOD_NOTE)}</div>
</div></section>"""


def headroom_view(h, ok_text):
    """(big number, unit, explanation) for a headroom result."""
    if h is None:
        return "n/a", "", ""
    if h["vms"] >= 0:
        return f"{h['vms']:,}", "more VMs", f"{ok_text} {constraint_name(h['limit'])} runs out first."
    return f"{-h['vms']:,}", "VMs short", "Already short of the design rules at this host count."


def headroom_cards_html(r):
    """'Room to grow' cards for a single-cluster report."""
    e = lambda v: escape(str(v))  # noqa: E731
    if r["headroom_now"] is None:
        return ""
    over = r["vcpu_ratio"] > r["hv_ratio"] + EPS
    note = (f" <span class='over'>The design limit ({r['vcpu_ratio']:.1f}:1) is above the {e(hv_label(r))}, "
            "so this headroom assumes a denser cluster than the guidance.</span>" if over else "")

    def card(cls, label, big, unit, sub):
        return (f"<div class='card {cls}'><div class='label'>{e(label)}</div><div class='big'>{e(big)} <small>{e(unit)}</small>"
                f"</div><p>{e(sub)}</p></div>")
    d1 = headroom_view(r["headroom_now"], "Before another host is needed.")
    fu = headroom_view(r["headroom_fut"], f"After {r['years']} years of {r['growth']*100:.0f}% growth.")
    cards = [
        card("navy", "Average VM", f"{r['avg_vm_vcpu']:.1f}", f"vCPU · {r['avg_vm_ram']:,.1f} GB",
             f"Across {r['tot_vms']:,} VMs in scope today."),
        card("green" if r["headroom_now"]["vms"] >= 0 else "", f"Day 1 · {r['hosts_now']} hosts", *d1),
        card("", f"Future · {r['hosts_fut']} hosts", *fu),
        card("", "Each extra host", f"~{r['vms_per_host']:,}" if r["vms_per_host"] else "n/a", "average VMs",
             "Room one more host adds at the design limit."),
    ]
    return f"""
  <h3>Room to grow</h3>
  <p class="foot" style="margin-top:-6px">How many more VMs of today's average size fit, with the HA hosts failed and the
    {r['vcpu_ratio']:.1f}:1 design limit, before another host is needed.{note}</p>
  <div class="cards">{''.join(cards)}</div>"""


# =============================================================================
# 3b. PROJECT LIBRARY (saved projects keep ONLY the columns the sizing needs)
# =============================================================================
PROJECT_FORMAT = "hostsizer-project"
PROJECT_VERSION = 3   # v2: settings per cluster scope; v3: several sources per project
PROJECT_EXT = ".hsproj"
PROJECTS_DIR = os.path.expanduser(os.environ.get("HOSTSIZER_PROJECTS", os.path.join("~", "HostSizer", "Projects")))


def _unit_cols(*bases):
    units = ("MiB", "MB", "GiB", "GB", "TiB", "TB")
    return [f"{b} {u}" for b in bases for u in units]


# Every column the parsers can read. Anything else in the export is dropped when a project is saved.
# Keep this in sync with parse_rvtools() / parse_live_optics(); save_project() verifies nothing is lost.
REQUIRED_COLUMNS = {
    "RVTools": {
        "vInfo": ["VM", "Powerstate", "Template", "SRM Placeholder", "Cluster", "CPUs", "Memory", "Memory GB", "Path"]
                 + _unit_cols("Provisioned", "In Use"),
        "vHost": ["Host", "Cluster", "# CPU", "Cores per CPU", "# Memory", "CPU Model", "CPU model",
                  "Speed", "CPU Speed", "Speed MHz", "ESX Version", "Product", "Vendor", "Model"],
        "vDatastore": ["Name", "Cluster name", "Cluster", "Type", "# Hosts"] + _unit_cols("Capacity", "In Use", "Provisioned"),
        "vPartition": ["VM"] + _unit_cols("Consumed"),
        "vDisk": RDM_DISK_COLS + _unit_cols("Capacity"),
        "vMultiPath": LUN_COLS,
        "vLicense": ["Name", "Used"],
        "vSnapshot": SNAP_COLS,
        "vMemory": ["VM", "Powerstate", "Reservation"] + _unit_cols("Size"),
        "vMetaData": ["xlsx creation datetime", "Server time", "xlsx creation date"],
    },
    "LiveOptics": {
        "VMs": ["VM Name", "Power State", "Template", "SRM Placeholder", "Cluster", "Virtual CPU",
                "Provisioned Memory (MiB)", "Virtual Disk Size (MiB)", "Virtual Disk Used (MiB)",
                "Guest VM Disk Used (MiB)"],
        "ESX Hosts": ["Cluster", "CPU Sockets", "CPU Cores", "Memory (KiB)", "CPU Model", "Processor", "CPU Type",
                      "CPU Description", "CPU Speed (MHz)", "Speed (MHz)", "CPU Clock (MHz)", "CPU Speed (GHz)",
                      "Speed (GHz)", "CPU Clock Speed (GHz)", "ESX Version", "OS Version", "Version", "OS",
                      "Manufacturer", "Model"] + HOST_NAME_COLS,
        "Host Devices": ["Cluster", "Capacity (GiB)", "Used Capacity (GiB)", "Free Capacity (GiB)", "Device Type",
                         "Vendor", "Model", "VM Count"] + HOST_NAME_COLS + LUN_ID_COLS,
        "Host Disks": ["Cluster", "Host", "Disk Name", "Disk Capacity (MiB)", "Disk Model", "Disk Vendor", "SSD"],
        "ESX Performance": ["Cluster", "95th Percentile CPU (GHz)", "Peak CPU (GHz)", "Average CPU (GHz)"],
        "ESX Licenses": ["Software Title"],
        SNAP_SHEET: SNAP_COLS,
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
    srcs = p.get("sources") or [p["source"]]
    types = sorted({x["type"] for x in srcs})
    files = srcs[0]["file_name"] + (f" + {len(srcs) - 1} more" if len(srcs) > 1 else "")
    return {k: p.get(k) for k in ("name", "saved", "created", "app_version", "summary")} | {
        "source_type": " + ".join(types), "file_name": files, "n_sources": len(srcs)}


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


# ---------- processor catalog (Target hardware > Processor) ----------
# Current and previous server generations, P-core / full-feature parts only. Base (rated) clock in GHz -
# guaranteed under full load, so it is the conservative figure for measured-CPU sizing.
# (model, cores per socket, base GHz, max sockets, note)
CPU_FAMILIES = [
    ("Intel", "Xeon 6 (Granite Rapids)", "Intel Xeon", [
        ("6505P", 12, 2.2, 2, ""), ("6507P", 8, 3.5, 2, ""), ("6511P", 16, 2.3, 1, ""), ("6515P", 16, 2.3, 2, ""),
        ("6517P", 16, 3.2, 2, ""), ("6520P", 24, 2.4, 2, ""), ("6521P", 24, 2.6, 1, ""), ("6527P", 24, 3.0, 2, ""),
        ("6530P", 32, 2.3, 2, ""), ("6714P", 8, 4.0, 8, ""), ("6724P", 16, 3.6, 8, ""), ("6728P", 24, 2.7, 8, ""),
        ("6730P", 32, 2.5, 2, ""), ("6731P", 32, 2.5, 1, ""), ("6732P", 32, 3.8, 2, ""), ("6736P", 36, 2.0, 2, ""),
        ("6737P", 32, 2.9, 2, ""), ("6738P", 32, 2.9, 8, ""), ("6740P", 48, 2.1, 2, ""), ("6741P", 48, 2.5, 1, ""),
        ("6745P", 32, 3.1, 2, ""), ("6747P", 48, 2.7, 2, ""), ("6748P", 48, 2.5, 2, ""), ("6760P", 64, 2.2, 2, ""),
        ("6761P", 64, 2.5, 1, ""), ("6762P", 64, 2.9, 2, ""), ("6767P", 64, 2.4, 2, ""), ("6768P", 64, 2.4, 8, ""),
        ("6774P", 64, 2.5, 1, ""), ("6776P", 64, 2.3, 2, ""), ("6781P", 80, 2.0, 1, ""), ("6787P", 86, 2.0, 2, ""),
        ("6788P", 86, 2.0, 8, ""), ("6944P", 72, 1.8, 2, "6900P"), ("6952P", 96, 2.1, 2, "6900P"),
        ("6960P", 72, 2.7, 2, "6900P"), ("6962P", 72, 2.7, 2, "6900P"), ("6966P-C", 96, 3.1, 2, "6900P"),
        ("6972P", 96, 2.4, 2, "6900P"), ("6979P", 120, 2.1, 2, "6900P"), ("6980P", 128, 2.0, 2, "6900P"),
    ]),
    ("Intel", "5th Gen Xeon Scalable (Emerald Rapids)", "Intel Xeon", [
        ("4509Y", 8, 2.6, 2, ""), ("4510", 12, 2.4, 2, ""), ("4514Y", 16, 2.0, 2, ""), ("4516Y+", 24, 2.2, 2, ""),
        ("5512U", 28, 2.1, 1, ""), ("5515+", 8, 3.2, 2, ""), ("5520+", 28, 2.2, 2, ""), ("6526Y", 16, 2.8, 2, ""),
        ("6530", 32, 2.1, 2, ""), ("6534", 8, 3.9, 2, ""), ("6538N", 32, 2.1, 2, ""), ("6538Y+", 32, 2.2, 2, ""),
        ("6542Y", 24, 2.9, 2, ""), ("6544Y", 16, 3.6, 2, ""), ("6548N", 32, 2.8, 2, ""), ("6548Y+", 32, 2.5, 2, ""),
        ("6554S", 36, 2.2, 2, ""), ("6558Q", 32, 3.2, 2, "liquid cooled"), ("8558", 48, 2.1, 2, ""),
        ("8558P", 48, 2.7, 2, ""), ("8558U", 48, 2.0, 1, ""), ("8562Y+", 32, 2.8, 2, ""), ("8568Y+", 48, 2.3, 2, ""),
        ("8570", 56, 2.1, 2, ""), ("8571N", 52, 2.4, 2, ""), ("8580", 60, 2.0, 2, ""), ("8581V", 60, 2.0, 1, ""),
        ("8592+", 64, 1.9, 2, ""), ("8592V", 64, 2.0, 2, ""), ("8593Q", 64, 2.2, 2, "liquid cooled"),
    ]),
    ("AMD", "EPYC 9005 (Turin)", "AMD EPYC", [
        ("9015", 8, 3.6, 2, ""), ("9115", 16, 2.6, 2, ""), ("9135", 16, 3.65, 2, ""), ("9175F", 16, 4.2, 2, ""),
        ("9255", 24, 3.25, 2, ""), ("9275F", 24, 4.1, 2, ""), ("9335", 32, 3.0, 2, ""), ("9355", 32, 3.55, 2, ""),
        ("9355P", 32, 3.55, 1, ""), ("9365", 36, 3.4, 2, ""), ("9375F", 32, 3.8, 2, ""), ("9455", 48, 3.15, 2, ""),
        ("9455P", 48, 3.15, 1, ""), ("9475F", 48, 3.65, 2, ""), ("9535", 64, 2.4, 2, ""), ("9555", 64, 3.2, 2, ""),
        ("9555P", 64, 3.2, 1, ""), ("9565", 72, 3.15, 2, ""), ("9575F", 64, 3.3, 2, ""), ("9655", 96, 2.5, 2, ""),
        ("9655P", 96, 2.5, 1, ""), ("9755", 128, 2.7, 2, ""), ("9645", 96, 2.3, 2, "Zen 5c dense"),
        ("9745", 128, 2.4, 2, "Zen 5c dense"), ("9825", 144, 2.2, 2, "Zen 5c dense"),
        ("9845", 160, 2.1, 2, "Zen 5c dense"), ("9965", 192, 2.25, 2, "Zen 5c dense"),
    ]),
    ("AMD", "EPYC 9004 (Genoa)", "AMD EPYC", [
        ("9124", 16, 3.0, 2, ""), ("9174F", 16, 4.1, 2, ""), ("9184X", 16, 3.55, 2, "3D V-Cache"),
        ("9224", 24, 2.5, 2, ""), ("9254", 24, 2.9, 2, ""), ("9274F", 24, 4.05, 2, ""), ("9334", 32, 2.7, 2, ""),
        ("9354", 32, 3.25, 2, ""), ("9354P", 32, 3.25, 1, ""), ("9374F", 32, 3.85, 2, ""),
        ("9384X", 32, 3.1, 2, "3D V-Cache"), ("9454", 48, 2.75, 2, ""), ("9454P", 48, 2.75, 1, ""),
        ("9474F", 48, 3.6, 2, ""), ("9534", 64, 2.45, 2, ""), ("9554", 64, 3.1, 2, ""), ("9554P", 64, 3.1, 1, ""),
        ("9634", 84, 2.25, 2, ""), ("9654", 96, 2.4, 2, ""), ("9654P", 96, 2.4, 1, ""),
        ("9684X", 96, 2.55, 2, "3D V-Cache"), ("9734", 112, 2.2, 2, "Zen 4c dense"),
        ("9754", 128, 2.25, 2, "Zen 4c dense"), ("9754S", 128, 2.25, 2, "Zen 4c dense, no SMT"),
    ]),
]
CPU_CUSTOM = "Custom (enter cores and GHz)"


def _cpu_catalog():
    """Flat processor list in dropdown order: vendor, generation (newest first), then cores and GHz."""
    out = []
    for vendor, gen, brand, models in CPU_FAMILIES:
        for model, cores, ghz, max_s, note in sorted(models, key=lambda m: (m[1], m[2], m[0])):
            extras = [x for x in (note, "1-socket only" if max_s == 1 else "") if x]
            label = (f"{vendor} {gen} · {model} · {cores} cores · {ghz:.2f} GHz"
                     + (f" · {', '.join(extras)}" if extras else ""))
            out.append({"label": label, "vendor": vendor, "gen": gen, "name": f"{brand} {model}",
                        "cores": cores, "ghz": ghz, "max_sockets": max_s, "note": note})
    return out


CPU_CATALOG = _cpu_catalog()
CPU_BY_LABEL = {c["label"]: c for c in CPU_CATALOG}
CPU_OPTIONS = [CPU_CUSTOM] + [c["label"] for c in CPU_CATALOG]
# RAM per host: typical 2-socket configurations on current servers (every channel filled, multiples of 256 GB)
RAM_OPTIONS = [256, 512, 768, 1024, 1536, 2048, 3072, 4096]
RAM_MIN, RAM_MAX = 64, 8192


def selected_cpu(p):
    """The chosen catalog processor, only while cores per socket and GHz still match it."""
    c = CPU_BY_LABEL.get(p.get("cpu_model") or "")
    if c and int(p.get("cores") or 0) == c["cores"] and round(float(p.get("clock") or 0), 2) == round(c["ghz"], 2):
        return c
    return None


def host_cpu_text(r):
    """'2 x Intel Xeon 6737P (32 cores, 2.90 GHz base)' or '2 sockets x 32 cores at 2.90 GHz'."""
    c = selected_cpu(r)
    if c:
        return f"{r['sockets']} x {c['name']} ({c['cores']} cores, {c['ghz']:.2f} GHz base)"
    return f"{r['sockets']} socket{'' if r['sockets'] == 1 else 's'} x {r['cores']} cores at {float(r['clock']):.2f} GHz"


def unit_mb(df, base):
    """Per-row decimal MB from a '{base} <unit>' column, or None when the export doesn't have it."""
    if df is None:
        return None
    for unit, f in UNIT_MB:
        col = f"{base} {unit}"
        if col in df.columns:
            return num(df[col]).fillna(0) * f
    return None


def mb_sum(df, base):
    s = unit_mb(df, base)
    return float(s.sum()) if s is not None else 0.0


def snapshot_size_mb(df):
    """Per-snapshot size in MB: the 'total' column when there is one (delta disks + memory), else any size column."""
    cols = list(df.columns)
    pick = None
    for pref in ("total", ""):
        for c in cols:
            lc = c.lower()
            if "size" in lc and "vmsn" not in lc and pref in lc and any(u.lower() in lc for u, _ in UNIT_MB):
                pick = c
                break
        if pick:
            break
    if not pick:
        return None
    f = next(f for u, f in sorted(UNIT_MB, key=lambda x: -len(x[0])) if u.lower() in pick.lower())
    return num(df[pick]).fillna(0) * f


def export_date(sheets):
    """When the export was taken (RVTools vMetaData), so snapshot ages aren't measured from today."""
    md = sheets.get("vMetaData")
    if md is not None and not md.empty:
        c = first_col(md, ["xlsx creation datetime", "Server time", "xlsx creation date"])
        if c:
            d = pd.to_datetime(md[c], errors="coerce").dropna()
            if not d.empty:
                return d.iloc[0].to_pydatetime().replace(tzinfo=None)
    return None


def snapshot_frame(sheet, vm_names, ref):
    """[vm, mb, days] for the snapshots of the VMs in scope."""
    if sheet is None or sheet.empty:
        return None
    vc = first_col(sheet, SNAP_VM_COLS)
    mb = snapshot_size_mb(sheet)
    if not vc or mb is None:
        return None
    df = pd.DataFrame({"vm": sheet[vc].astype(str), "mb": mb})
    dc = first_col(sheet, SNAP_DATE_COLS)
    when = pd.to_datetime(sheet[dc], errors="coerce") if dc else pd.Series(pd.NaT, index=sheet.index)
    try:
        when = when.dt.tz_localize(None)
    except (TypeError, AttributeError):
        pass
    df["days"] = (pd.Timestamp(ref or datetime.now()) - when).dt.days
    return df[df["vm"].isin(vm_names)].reset_index(drop=True)


def add_snapshots(sto, snaps, ref):
    if snaps is None:
        sto["notes"].append("Snapshot data not in this export (RVTools vSnapshot tab or a Live Optics snapshot tab)")
        return
    sto.update(has_snap=True, snap_mb=float(snaps["mb"].sum()), snap_cnt=len(snaps), snap_vms=int(snaps["vm"].nunique()),
               snap_dated=bool(snaps["days"].notna().any()), snap_ref=ref.strftime("%b %d, %Y") if ref else None)


def storage_rvtools(sheets, info_scope, vms_info, clusters, power, ds):
    """RVTools storage: VM files, snapshots, swap, templates, RDMs, guest used and other datastore files."""
    sto = copy.deepcopy(EMPTY_STO)
    names = set(vms_info["VM"].astype(str)) if "VM" in vms_info.columns else set()
    sto["vm_alloc"], sto["vm_inuse"] = mb_sum(vms_info, "Provisioned"), mb_sum(vms_info, "In Use")
    if "Template" in info_scope.columns and power == POWER_ALL:   # templates move with the estate
        t = info_scope[is_true(info_scope["Template"])]
        sto.update(tmpl_cnt=len(t), tmpl_alloc=mb_sum(t, "Provisioned"), tmpl_inuse=mb_sum(t, "In Use"))
    part = sheets.get("vPartition")
    if part is not None and "VM" in part.columns:
        sto["guest"] = mb_sum(part[part["VM"].astype(str).isin(names)], "Consumed")
    disk = sheets.get("vDisk")
    if disk is not None and "Raw" in disk.columns:
        d = disk[disk["VM"].astype(str).isin(names)] if "VM" in disk.columns else disk
        rdm = d[is_true(d["Raw"])]
        sto.update(rdm_cnt=len(rdm), rdm=mb_sum(rdm, "Capacity"))
    sto["rdms"] = rdm_frame(disk, names)
    sto["host_luns"] = host_lun_frame(sheets.get("vMultiPath"), clusters, sto["rdms"])
    mem = sheets.get("vMemory")
    if mem is not None and "VM" in mem.columns and unit_mb(mem, "Size") is not None:
        m = mem[mem["VM"].astype(str).isin(names)]
        if "Powerstate" in m.columns:
            m = m[m["Powerstate"].astype(str).str.contains("poweredOn", case=False, na=False)]
        res = num_col(m, "Reservation").fillna(0) * MIB_MB
        sto.update(swap=float((unit_mb(m, "Size") - res).clip(lower=0).sum()), has_swap=True)
    # vSAN today: VM usage on vSAN datastores includes the policy's copies (removed in storage_summary)
    dsall = sheets.get("vDatastore")
    if ds.get("df") is not None and "Type" in ds["df"].columns:
        v = ds["df"][ds["df"]["Type"].astype(str).str.contains("vsan", case=False, na=False)]
        sto.update(vsan_ds=len(v), vsan_cap=mb_sum(v, "Capacity"), vsan_used=mb_sum(v, "In Use"))
    if dsall is not None and {"Name", "Type"} <= set(dsall.columns):
        vsan_ds = set(dsall.loc[dsall["Type"].astype(str).str.contains("vsan", case=False, na=False), "Name"].astype(str))
        if vsan_ds and "Path" in info_scope.columns:
            home = lambda df: df["Path"].astype(str).str.extract(r"^\[([^\]]+)\]")[0].isin(vsan_ds)  # noqa: E731
            on_v = vms_info[home(vms_info)]
            sto.update(vsan_src=bool(len(on_v)) or bool(ds.get("vsan")), vsan_vm_alloc=mb_sum(on_v, "Provisioned"),
                       vsan_vm_inuse=mb_sum(on_v, "In Use"), vsan_vms=len(on_v))
            if "Template" in info_scope.columns and power == POWER_ALL:
                t = info_scope[is_true(info_scope["Template"])]
                t = t[home(t)]
                sto.update(vsan_tmpl_alloc=mb_sum(t, "Provisioned"), vsan_tmpl_inuse=mb_sum(t, "In Use"))
        elif ds.get("vsan"):
            sto["vsan_src"] = True
            sto["notes"].append("vSAN datastores found but VM paths aren't in the export, so today's vSAN protection "
                                "copies can't be removed from VM usage")
    ref = export_date(sheets)
    snaps = snapshot_frame(sheets.get("vSnapshot"), names, ref)
    add_snapshots(sto, snaps, ref)
    sto["snaps"] = snaps
    sto.update(ds_cap=ds["cap_mb"], ds_used=ds["used_mb"], ds_alloc=ds["alloc_mb"])
    if ds["cap_mb"] > 0 and power == POWER_ALL and ds["mapped"]:
        if ds["vsan"]:
            sto["notes"].append("Other datastore files not estimated: vSAN datastore usage includes its protection copies")
        else:
            sto["other"] = other_files_mb(sheets["vInfo"], ds["df"], ds.get("ccol"), clusters)
            if sto["other"] is None:
                sto["notes"].append("Other datastore files not estimated: VM paths (vInfo 'Path') not in export")
    return sto


def other_files_mb(vinfo, dstores, ccol, clusters):
    """Datastore space not used by any VM or template (ISOs, orphaned and non-VM files). Each VM's usage is
    attributed to its home datastore (vInfo 'Path'); a datastore shared by several clusters is split evenly so
    cluster totals don't count it twice."""
    use = unit_mb(vinfo, "In Use")
    if "Path" not in vinfo.columns or use is None or "Name" not in dstores.columns:
        return None
    home = vinfo["Path"].astype(str).str.extract(r"^\[([^\]]+)\]")[0]
    by_ds = use.groupby(home).sum()
    used = unit_mb(dstores, "In Use")
    if used is None:
        return None
    wanted = set(clusters or [])
    total = 0.0
    for idx, row in dstores.iterrows():
        other = max(float(used.loc[idx]) - float(by_ds.get(str(row["Name"]), 0.0)), 0.0)
        share = 1.0
        if clusters is not None and ccol:
            names = [c.strip() for c in str(row[ccol]).split(",") if c.strip()]
            share = sum(c in wanted for c in names) / len(names) if names else 1.0
        total += other * share
    return total


BOOT_DISK_WORDS = ("boss", "satadom", "usb", "sd card", "m.2")


def lo_vsan(sheets, esx, vms_df):
    """Live Optics has no datastore type, so vSAN is recognised from the hardware: a shared ('Cluster') datastore
    whose vendor/model is one of the hosts' local disks, or VxRail hosts."""
    hc = first_col(esx, HOST_NAME_COLS)
    hosts = set(esx[hc].astype(str)) if hc else set()
    vxrail = "Model" in esx.columns and esx["Model"].astype(str).str.contains("vxrail", case=False, na=False).any()
    disks = sheets.get("Host Disks")
    local = pd.DataFrame()
    if disks is not None and {"Host", "Disk Name"} <= set(disks.columns):
        d = disks[disks["Host"].astype(str).isin(hosts)]
        model = d["Disk Model"].astype(str) if "Disk Model" in d.columns else pd.Series("", index=d.index)
        txt = (d["Disk Name"].astype(str) + " " + model).str.lower()
        local = d[d["Disk Name"].astype(str).str.lower().str.startswith("local")
                  & ~txt.str.contains("|".join(BOOT_DISK_WORDS), na=False)]
    dev = sheets.get("Host Devices")
    found = pd.DataFrame()
    if dev is not None and {"Device Type", "Vendor", "Model"} <= set(dev.columns):
        sc = first_col(dev, HOST_NAME_COLS)
        dv = dev[dev[sc].astype(str).isin(hosts)] if sc else dev
        shared = dv[dv["Device Type"].astype(str).str.lower() == "cluster"]
        idc = first_col(shared, LUN_ID_COLS)
        if idc:
            shared = shared.drop_duplicates(subset=[idc])
        if len(local) and {"Disk Vendor", "Disk Model"} <= set(local.columns):
            pairs = set(zip(local["Disk Vendor"].astype(str).str.lower(), local["Disk Model"].astype(str).str.lower()))
            found = shared[[(str(v).lower(), str(m).lower()) in pairs for v, m in zip(shared["Vendor"], shared["Model"])]]
        if found.empty and vxrail and len(shared) == 1:
            found = shared
    if found.empty and not vxrail:
        return {}
    out = {"vsan_src": True, "vsan_ds": len(found), "vsan_vxrail": bool(vxrail), "vsan_copies_in_vm": False,
           "vsan_cap": col_sum(found, "Capacity (GiB)") * GIB_MB, "vsan_used": col_sum(found, "Used Capacity (GiB)") * GIB_MB}
    names = set(found[first_col(found, LUN_ID_COLS)].astype(str)) if len(found) and first_col(found, LUN_ID_COLS) else set()
    if "Datastore" in vms_df.columns and names:
        on = vms_df[vms_df["Datastore"].astype(str).isin(names)]
        out["vsan_vms"] = len(on)
        out["vsan_vm_used_logical"] = col_sum(on, "Virtual Disk Used (MiB)") * MIB_MB
    if len(local) and "Disk Capacity (MiB)" in local.columns:
        n_hosts = max(local["Host"].nunique(), 1)
        tiers = (local.assign(tb=(num(local["Disk Capacity (MiB)"]) * MIB_MB / 1e6).round(2))
                 .groupby(["tb", local.get("Disk Model", pd.Series("", index=local.index)).astype(str)]).size())
        parts = [f"{cnt / n_hosts:g} x {tb:g} TB {model}".strip() for (tb, model), cnt in tiers.sort_index(ascending=False).items()]
        out["vsan_layout"] = f"{n_hosts} hosts, per host: " + ", ".join(parts)
        out["vsan_kind"] = ("OSA with cache and capacity tiers" if len(tiers) == 2 and tiers.index.get_level_values(0).min()
                            < tiers.index.get_level_values(0).max() / 4 else "single tier (ESA, or OSA all-flash)")
        out["vsan_raw"] = float((num(local["Disk Capacity (MiB)"]) * MIB_MB).sum())
    return out


def storage_liveoptics(sheets, vm_scope, vms_df, power, esx=None):
    sto = copy.deepcopy(EMPTY_STO)
    mib = lambda df, c: col_sum(df, c) * MIB_MB  # noqa: E731
    vc = "VM Name" if "VM Name" in vms_df.columns else None
    names = set(vms_df[vc].astype(str)) if vc else set()
    sto.update(vm_alloc=mib(vms_df, "Virtual Disk Size (MiB)"), vm_inuse=mib(vms_df, "Virtual Disk Used (MiB)"),
               guest=mib(vms_df, "Guest VM Disk Used (MiB)"))
    if "Template" in vm_scope.columns and power == POWER_ALL:
        t = vm_scope[is_true(vm_scope["Template"])]
        sto.update(tmpl_cnt=len(t), tmpl_alloc=mib(t, "Virtual Disk Size (MiB)"), tmpl_inuse=mib(t, "Virtual Disk Used (MiB)"))
    snaps = snapshot_frame(sheets.get(SNAP_SHEET), names, None)
    add_snapshots(sto, snaps, None)
    if snaps is not None and len(snaps):
        sto["notes"].append("Live Optics snapshot ages are measured from today, not the capture date")
    sto["snaps"] = snaps
    sto["notes"].append("Live Optics: other datastore files can't be separated from LUN-level capacity")
    if esx is not None and not esx.empty:
        sto.update(lo_vsan(sheets, esx, vms_df))
    return sto


def combine_storage(stos):
    out = copy.deepcopy(EMPTY_STO)
    for s in stos:
        for k, v in s.items():
            if k in ("notes", "snaps", "snap_ref", "other", "has_snap", "has_swap", "snap_dated", "vsan_src",
                     "rdms", "host_luns", "vsan_layout", "vsan_kind", "vsan_vxrail", "vsan_copies_in_vm"):
                continue
            out[k] += v
    others = [s["other"] for s in stos if s.get("other") is not None]
    out["other"] = sum(others) if others else None
    out["has_snap"] = any(s["has_snap"] for s in stos)
    out["has_swap"] = any(s["has_swap"] for s in stos)
    out["vsan_src"] = any(s.get("vsan_src") for s in stos)
    out["vsan_vxrail"] = any(s.get("vsan_vxrail") for s in stos)
    out["vsan_copies_in_vm"] = any(s.get("vsan_copies_in_vm") and s.get("vsan_src") for s in stos)
    out["vsan_layout"] = " | ".join(dict.fromkeys(s["vsan_layout"] for s in stos if s.get("vsan_layout")))
    out["vsan_kind"] = " | ".join(dict.fromkeys(s["vsan_kind"] for s in stos if s.get("vsan_kind")))
    out["snap_dated"] = any(s["snap_dated"] for s in stos)
    out["snap_ref"] = next((s["snap_ref"] for s in stos if s.get("snap_ref")), None)
    out["notes"] = list(dict.fromkeys(n for s in stos for n in s["notes"]))
    for k in ("rdms", "host_luns"):
        fr = [s[k] for s in stos if s.get(k) is not None]
        out[k] = pd.concat(fr, ignore_index=True) if fr else None
    frames = [s["snaps"] for s in stos if s.get("snaps") is not None]
    out["snaps"] = pd.concat(frames, ignore_index=True) if frames else None
    return out


def storage_summary(sto, p, mult):
    """Consumption totals and the effective capacity to buy: consumed x growth / (1 - free-space target)."""
    s = {k: v for k, v in sto.items() if k != "snaps"}
    old_days = int(p.get("snap_old_days") or DEFAULTS["snap_old_days"])
    free = min(max(float(p.get("sto_free_pct", DEFAULTS["sto_free_pct"])), 0.0), 90.0) / 100
    snaps = sto.get("snaps")
    old = snaps[snaps["days"] > old_days] if snaps is not None and len(snaps) else None
    s.update(old_days=old_days, free_pct=free * 100, mult=mult,
             snap_old_cnt=len(old) if old is not None else 0,
             snap_old_mb=float(old["mb"].sum()) if old is not None else 0.0,
             snap_oldest=int(snaps["days"].max()) if snaps is not None and snaps["days"].notna().any() else None)
    if s.get("vsan_src") and s.get("vsan_copies_in_vm"):   # RVTools: VM usage on vSAN includes protection copies
        f = VSAN_SRC_POLICIES.get(VSAN_SRC_LEGACY.get(p.get("vsan_src_policy"), p.get("vsan_src_policy")), 2.0)
        cut = 1 - 1 / f
        s["vsan_removed"] = (s["vsan_vm_inuse"] + s["vsan_tmpl_inuse"]) * cut
        for k in ("vm_alloc", "vm_inuse", "tmpl_alloc", "tmpl_inuse"):
            s[k] -= s[f"vsan_{k}"] * cut
    other = s["other"] or 0.0
    snap, snap_old = s["snap_mb"], s["snap_old_mb"]
    tot = {"inuse": s["vm_inuse"] + s["tmpl_inuse"] + s["rdm"] + other,
           "alloc": s["vm_alloc"] + s["tmpl_alloc"] + s["rdm"] + other}
    s["totals"] = {b: {"with": v, "no_old": max(v - snap_old, 0.0), "without": max(v - snap, 0.0)} for b, v in tot.items()}
    s["required"] = {b: {k: (v / (1 - free), v * mult / (1 - free)) for k, v in row.items()}
                     for b, row in s["totals"].items()}
    return s


def fmt_cap(mb):
    """Decimal TB (2 decimals); MB only for amounts under 0.01 TB."""
    if mb is None:
        return "n/a"
    return f"{mb / 1e6:,.2f} TB" if abs(mb) >= 1e4 or mb == 0 else f"{mb:,.0f} MB"


STO_COLS = [("with", "With all snapshots"), ("no_old", "Without old snapshots"), ("without", "Without snapshots")]


def storage_tables_html(r, table_class="compare", hl_class="hl"):
    """Consumption and effective-capacity tables, shared by the app and the reports."""
    s = r["sto"]
    e = lambda v: escape(str(v))  # noqa: E731
    snap_old = f"older than {s['old_days']} days"
    snaps = s["has_snap"] and s["snap_cnt"] > 0
    rows = [(f"Virtual machines ({r['tot_vms']:,})", s["vm_alloc"] - s["snap_mb"], s["vm_inuse"] - s["snap_mb"], "")]
    if snaps:
        rows.append((f"Snapshots ({s['snap_cnt']:,})", s["snap_mb"], s["snap_mb"], ""))
        if s["snap_dated"] and s["snap_old_cnt"]:
            rows.append((f"&nbsp;&nbsp;of which {snap_old} ({s['snap_old_cnt']:,})", s["snap_old_mb"], s["snap_old_mb"], "sub"))
    if s["tmpl_cnt"]:
        rows.append((f"Templates ({s['tmpl_cnt']:,})", s["tmpl_alloc"], s["tmpl_inuse"], ""))
    if s["rdm_cnt"]:
        rows.append((f"RDM disks ({s['rdm_cnt']:,})", s["rdm"], s["rdm"], ""))
    if s["other"]:
        rows.append(("Other datastore files (ISOs and non-VM files, estimate)", None, s["other"], ""))
    body = ""
    for label, a, u, cls in rows:
        lab = label if label.startswith("&nbsp;") else e(label)
        tr = "<tr class='sub'>" if cls else "<tr>"
        body += (f"{tr}<td>{lab}</td><td>{fmt_cap(a) if a is not None else '-'}</td>"
                 f"<td>{fmt_cap(u) if u is not None else 'n/a'}</td></tr>")
    t = s["totals"]
    body += (f"<tr class='total'><td>Total{' with snapshots' if snaps else ''}</td><td>{fmt_cap(t['alloc']['with'])}</td>"
             f"<td class='{hl_class}'>{fmt_cap(t['inuse']['with'])}</td></tr>")
    if snaps:
        body += (f"<tr class='total'><td>Total without snapshots</td><td>{fmt_cap(t['alloc']['without'])}</td>"
                 f"<td>{fmt_cap(t['inuse']['without'])}</td></tr>")
    consumed = (f"<table class='{table_class}'><tr><th>Consumed today</th><th>Allocated (provisioned)</th>"
                f"<th>In use</th></tr>{body}</table>")
    req = s["required"]
    cols = [c for c in STO_COLS if snaps and (c[0] != "no_old" or s["snap_old_cnt"])] or [("with", "Required")]
    rbody = ""
    for b, name in (("inuse", "In use"), ("alloc", "Allocated")):
        for i, when in ((0, "Day 1"), (1, f"Future · {r['years']} yr")):
            cells = ""
            for k, _ in cols:
                td = f"<td class='{hl_class}'>" if (b, i, k) == ("inuse", 1, "with") else "<td>"
                cells += f"{td}{fmt_cap(req[b][k][i])}</td>"
            rbody += f"<tr><td>{name} · {when}</td>{cells}</tr>"
    heads = "".join(f"<th>{e(f'Without snapshots {snap_old}' if k == 'no_old' else h)}</th>" for k, h in cols)
    required = f"<table class='{table_class}'><tr><th>Effective capacity required</th>{heads}</tr>{rbody}</table>"
    return consumed, required


def storage_facts(r):
    """Reference figures shown under the tables: (label, value)."""
    s = r["sto"]
    out = [("Used inside the guest operating systems", fmt_cap(s["guest"]))]
    if s["has_swap"]:
        out.append(("VM swap files (powered-on VMs, memory minus reservation; part of VM in use)", fmt_cap(s["swap"])))
    if s["has_snap"]:
        out.append(("Oldest snapshot", f"{s['snap_oldest']:,} days" if s["snap_oldest"] is not None else "date not in export"))
    if not s["has_snap"]:
        out.append(("Snapshots", "Not in this export"))
    if s.get("vsan_layout"):
        out.append((f"Current vSAN drives ({s.get('vsan_kind')})", s["vsan_layout"]))
    if s.get("vsan_src") and s.get("vsan_copies_in_vm"):
        out.append((f"vSAN today: protection copies removed from {s['vsan_vms']:,} VMs ({src_policy_label(r)})",
                    fmt_cap(s.get("vsan_removed", 0.0))))
    out.append(("Growth and free space", f"{r['growth'] * 100:.0f}% a year for {r['years']} years (x{s['mult']:.2f}); "
                                         f"{s['free_pct']:.0f}% kept free"))
    return out


# ---------- vSAN ESA host storage estimate (3.13) ----------
# Model calibrated to Broadcom's vSAN ReadyNode Sizer (ESA, all flash): 6 hosts x 12 x 1.6 TB, RAID-6,
# Operations Reserve on, compression 1.2 -> 51.87 TB usable. Planning estimate only; confirm in the sizer.
VSAN_DRIVES_TB = [1.6, 1.92, 3.2, 3.84, 6.4, 7.68, 12.8, 15.36]
VSAN_MIN_DRIVES = 2
VSAN_FIXED = [("Fixed metadata overhead", 0.10), ("File system overhead", 0.05), ("Performance tier overhead", 0.05)]
VSAN_OBJECT_OVERHEAD = 0.131     # object/namespace overhead as a share of space left after reserves
VSAN_OR_RAW_FACTOR = 1.10        # Operations Reserve is usable space; raw = OR x protection x 1.10
VSAN_RESERVE_CAP = 0.30          # reserves capped below 30% of capacity for small clusters
VSAN_POLICIES = {                # name: (capacity multiplier, minimum hosts)
    "RAID-1 (FTT=1)": (2.0, 2),
    "RAID-5 (FTT=1, 2+1)": (1.5, 3),
    "RAID-5 (FTT=1, 4+1)": (1.25, 6),
    "RAID-6 (FTT=2, 4+2)": (1.5, 6),
}                                # vSAN ESA policies only (OSA isn't modelled for new clusters)
VSAN_AUTO = "Auto (by host count)"
VSAN_PICK_AUTO = "Closest fit (automatic)"
# Today's vSAN policy, used only to back the protection copies out of RVTools VM usage (the source cluster
# may be OSA or ESA). Multiplier = raw space per GB of VM data.
VSAN_SRC_POLICIES = {
    "RAID-1 mirror, FTT=1 (OSA or ESA) · 2x": 2.0,
    "RAID-5 3+1, FTT=1 (OSA) · 1.33x": 4 / 3,
    "RAID-5 2+1, FTT=1 (ESA, 3-5 hosts) · 1.5x": 1.5,
    "RAID-5 4+1, FTT=1 (ESA, 6+ hosts) · 1.25x": 1.25,
    "RAID-6 4+2, FTT=2 (OSA or ESA) · 1.5x": 1.5,
    "RAID-1 mirror, FTT=2 · 3x": 3.0,
    "No adjustment (use VM usage as reported) · 1x": 1.0,
}
VSAN_SRC_LEGACY = {"RAID-1 (FTT=1) · 2x": "RAID-1 mirror, FTT=1 (OSA or ESA) · 2x",
                   "RAID-5 · 1.33x": "RAID-5 3+1, FTT=1 (OSA) · 1.33x",
                   "RAID-5 (ESA 2+1) · 1.5x": "RAID-5 2+1, FTT=1 (ESA, 3-5 hosts) · 1.5x",
                   "RAID-6 · 1.5x": "RAID-6 4+2, FTT=2 (OSA or ESA) · 1.5x",
                   "RAID-1 (FTT=2) · 3x": "RAID-1 mirror, FTT=2 · 3x",
                   "Already excludes copies · 1x": "No adjustment (use VM usage as reported) · 1x"}


def vsan_policy(hosts, choice=VSAN_AUTO):
    """Protection by cluster size: 2 hosts RAID-1, 3-5 RAID-5, 6+ RAID-6 (unless a policy is chosen)."""
    if choice in VSAN_POLICIES:
        return choice
    if hosts <= 2:
        return "RAID-1 (FTT=1)"
    return "RAID-5 (FTT=1, 2+1)" if hosts <= 5 else "RAID-6 (FTT=2, 4+2)"


def vsan_capacity(hosts, drives, size_tb, policy, compression=1.0):
    """Capacity breakdown in TB for one host configuration. Operations Reserve from 3 hosts, Host Rebuild
    Reserve (one host's share, N+1) from 4 hosts; neither applies to 2-host clusters."""
    mult = VSAN_POLICIES[policy][0]
    raw = hosts * drives * size_tb
    parts = [(name, raw * pct) for name, pct in VSAN_FIXED]
    after = raw - sum(v for _, v in parts)
    gb = size_tb * 1000
    or_on, hrr_on = hosts >= 3, hosts >= 4
    op = ((hosts * 765 + hosts * drives * (min(gb * 0.05, 100) + min(gb * 0.0025, 100))) / 1000
          * mult * VSAN_OR_RAW_FACTOR) if or_on else 0.0
    hrr = after / hosts if hrr_on else 0.0
    if hrr_on and op + hrr > after * VSAN_RESERVE_CAP:
        hrr = max(after * VSAN_RESERVE_CAP - op, 0.0)
    left = max(after - op - hrr, 0.0)
    obj = left * VSAN_OBJECT_OVERHEAD
    data = left - obj
    usable = data / mult
    return {"hosts": hosts, "drives": drives, "size": size_tb, "policy": policy, "mult": mult, "raw": raw,
            "parts": parts, "op": op, "hrr": hrr, "or_on": or_on, "hrr_on": hrr_on, "obj": obj,
            "raid": data - usable, "usable": usable, "compression": compression, "effective": usable * compression}


def vsan_fit(hosts, target_tb, size_tb, policy, compression, max_drives):
    """Fewest drives per host of one size that reach the target, or None when even max_drives is short."""
    for n in range(VSAN_MIN_DRIVES, int(max_drives) + 1):
        c = vsan_capacity(hosts, n, size_tb, policy, compression)
        if c["effective"] >= target_tb - 1e-9:
            return c
    return None


def vsan_estimate(r):
    """Drive options for the Day 1 host count holding the future need (plus Day 1 need and future hosts)."""
    s = r["sto"]
    hosts, fut_hosts = int(r["hosts_now"]), int(r["hosts_fut"])
    comp = max(float(r.get("vsan_compression") or 1.0), 1.0)
    max_d = int(r.get("vsan_max_drives") or DEFAULTS["vsan_max_drives"])
    choice = r.get("vsan_raid") or VSAN_AUTO
    policy, fut_policy = vsan_policy(hosts, choice), vsan_policy(fut_hosts, choice)
    if VSAN_POLICIES[policy][1] > hosts:
        return {"error": f"{policy} needs at least {VSAN_POLICIES[policy][1]} hosts; the cluster has {hosts}."}
    if hosts < 2:
        return {"error": "vSAN needs at least 2 hosts."}
    # Reserves give vSAN its slack, so the target is consumption x growth without the free-space target;
    # 2-host clusters have no reserves, so they keep the free-space target.
    free = 0.0 if hosts >= 3 else s["free_pct"] / 100
    base = s["totals"]["inuse"]["with"] / 1e6
    tgt_fut, tgt_now = base * s["mult"] / (1 - free), base / (1 - free)
    fut_free = 0.0 if fut_hosts >= 3 else s["free_pct"] / 100
    tgt_fut_hosts = base * s["mult"] / (1 - fut_free)
    rows = []
    for size in VSAN_DRIVES_TB:
        rows.append({"size": size, "fit": vsan_fit(hosts, tgt_fut, size, policy, comp, max_d),
                     "now": vsan_fit(hosts, tgt_now, size, policy, comp, max_d),
                     "fut_hosts": vsan_fit(fut_hosts, tgt_fut_hosts, size, fut_policy, comp, max_d)
                     if VSAN_POLICIES[fut_policy][1] <= fut_hosts else None})
    fits = [x for x in rows if x["fit"]]
    best = min(fits, key=lambda x: (x["fit"]["raw"], x["fit"]["drives"])) if fits else None
    fewest = min(fits, key=lambda x: (x["fit"]["drives"], x["fit"]["raw"])) if fits else None
    return {"hosts": hosts, "fut_hosts": fut_hosts, "policy": policy, "fut_policy": fut_policy, "compression": comp,
            "max_drives": max_d, "target_fut": tgt_fut, "target_now": tgt_now, "target_fut_hosts": tgt_fut_hosts,
            "free_used": free, "rows": rows, "best": best, "fewest": fewest,
            "chosen": next((x for x in fits if vsan_pick_label(x) == r.get("vsan_pick")), best),
            "picked": any(vsan_pick_label(x) == r.get("vsan_pick") for x in fits)}


def vsan_pick_label(row):
    """Saved choice for one drive option: the drive size (so it survives changes in drive count)."""
    return f"{row['size']:g} TB"


def vsan_pick_text(v, value):
    row = next((x for x in v.get("rows", []) if x["fit"] and vsan_pick_label(x) == value), None)
    if row:
        return f"{row['fit']['drives']} x {row['size']:g} TB per host"
    b = v.get("best")
    return f"{VSAN_PICK_AUTO} · {b['fit']['drives']} x {b['size']:g} TB" if b else VSAN_PICK_AUTO


def vsan_pick_options(v):
    return [VSAN_PICK_AUTO] + [vsan_pick_label(x) for x in v.get("rows", []) if x["fit"]]


def vsan_tables_html(r, v, table_class="compare sto", hl_class="hl"):
    """Drive options table and the capacity breakdown of the closest fit (shared by app and reports)."""
    e = lambda x: escape(str(x))  # noqa: E731
    tb = lambda x: f"{x:,.2f} TB"  # noqa: E731
    body = ""
    for row in v["rows"]:
        f = row["fit"]
        tags = []
        if v["chosen"] is row:
            tags.append("In report")
        if v["best"] is row:
            tags.append("Closest fit")
        if v["fewest"] is row and v["fewest"] is not v["best"]:
            tags.append("Fewest drives")
        tag = "".join(f" <span class='tag mini'>{e(t)}</span>" for t in tags)
        cell = lambda c: f"{c['drives']}" if c else f"&gt; {v['max_drives']}"  # noqa: E731
        hl = f" class='{hl_class}'" if v["chosen"] is row else ""
        body += (f"<tr><td>{row['size']:g} TB NVMe{tag}</td><td{hl}>{cell(f)}</td>"
                 f"<td>{tb(f['raw']) if f else '-'}</td><td>{tb(f['effective']) if f else '-'}</td>"
                 f"<td>{cell(row['now'])}</td><td>{cell(row['fut_hosts']) if v['fut_policy'] else '-'}</td></tr>")
    options = (f"<table class='{table_class}'><tr><th>Drive</th><th>Per host</th><th>Raw</th><th>Effective</th>"
               f"<th>Day 1 need</th><th>{v['fut_hosts']} hosts</th></tr>{body}</table>")
    if not v["chosen"]:
        return options, ""
    c = v["chosen"]["fit"]
    rows = [("Raw capacity", c["raw"])] + c["parts"]
    rows += [("Operations reserve", c["op"]) if c["or_on"] else ("Operations reserve (not on 2-host clusters)", 0.0),
             ("Host rebuild reserve (N+1)", c["hrr"]) if c["hrr_on"] else ("Host rebuild reserve (needs 4+ hosts)", 0.0),
             ("Object overhead", c["obj"]), (f"Protection overhead · {c['policy']}", c["raid"]),
             ("Usable capacity", c["usable"])]
    if c["compression"] > 1:
        rows.append((f"Effective with {c['compression']:.2f}x compression", c["effective"]))
    bd = ""
    for i, (k, val) in enumerate(rows):
        tr = "<tr class='total'>" if i in (0, len(rows) - 1) else "<tr>"
        bd += f"{tr}<td>{e(k)}</td><td>{tb(val)}</td><td>{val / c['raw']:.0%}</td></tr>"
    breakdown = (f"<table class='{table_class}'><tr><th>{c['hosts']} hosts x {c['drives']} x "
                 f"{c['size']:g} TB</th><th>Capacity</th><th>Share of raw</th></tr>{bd}</table>")
    return options, breakdown


def src_policy_label(r):
    v = r.get("vsan_src_policy")
    return VSAN_SRC_LEGACY.get(v, v)


def vsan_detect_text(r):
    s = r["sto"]
    n = s.get("vsan_ds", 0)
    bits = [f"{n} datastore{'s' if n != 1 else ''}" if n else "VxRail hosts"]
    if s.get("vsan_cap"):
        bits.append(f"{fmt_cap(s['vsan_cap'])} capacity, {fmt_cap(s.get('vsan_used', 0))} used")
    if s.get("vsan_vms"):
        bits.append(f"{s['vsan_vms']:,} VMs")
    text = "vSAN" + (" (VxRail)" if s.get("vsan_vxrail") else "") + " in the current environment: " + ", ".join(bits) + "."
    logical = s.get("vsan_vm_used_logical") or 0
    if logical and s.get("vsan_used"):
        ratio = s["vsan_used"] / logical
        hint = "RAID-1 mirroring" if ratio >= 1.8 else ("RAID-5/6" if ratio >= 1.15 else "little protection overhead")
        text += f" Datastore used is {ratio:.1f}x the VMs' data, consistent with {hint}."
    if s.get("vsan_copies_in_vm"):
        text += f" VM usage is counted without today's protection copies ({src_policy_label(r)})."
    return text


def vsan_options_note(v):
    return (f"Per host: capacity drives per host on {v['hosts']} hosts for the future need, with the raw and effective "
            f"capacity that gives. Day 1 need: drives for today's requirement. {v['fut_hosts']} hosts: drives if the "
            "cluster grows to the future host count. Closest fit = least raw capacity that meets the need.")


def vsan_lede(r, v):
    if v.get("error"):
        return v["error"]
    src = (" The current environment runs on vSAN: VM usage there includes today's protection copies, which are "
           f"removed using the '{src_policy_label(r)}' assumption."
           if r["sto"].get("vsan_src") and r["sto"].get("vsan_copies_in_vm") else
           (" The current environment runs on vSAN; Live Optics VM usage already excludes its protection copies."
            if r["sto"].get("vsan_src") else ""))
    best = v["chosen"]
    fit = (f"The proposed configuration is {best['fit']['drives']} x {best['size']:g} TB NVMe capacity drives per host "
           f"({best['fit']['effective']:,.2f} TB effective)." if best else
           f"No single drive size reaches it within {v['max_drives']} drives per host; add hosts or raise the limit.")
    return (f"{v['hosts']} vSAN ESA hosts with {v['policy']} need {v['target_fut']:,.2f} TB of effective capacity for the "
            f"future workload (in use with snapshots, {r['years']} years of growth"
            + (f", {v['free_used'] * 100:.0f}% free" if v["free_used"] else ", reserves instead of a free-space target")
            + f"). {fit}{src}")


VSAN_NOTE = ("vSAN ESA estimate, all flash: 20% fixed overhead (metadata 10%, file system 5%, performance tier 5%), "
             "Broadcom's ESA Operations Reserve formula (3+ hosts), Host Rebuild Reserve of one host's share (4+ hosts, "
             "reserves capped at 30%), about 13% object overhead, then the protection policy and compression. "
             "Policy by host count: 2 hosts RAID-1, 3-5 RAID-5 (2+1), 6+ RAID-6 (4+2). Calibrated to the vSAN "
             "ReadyNode Sizer; confirm the final design there.")


def vsan_section_html(r, cls=""):
    e = lambda x: escape(str(x))  # noqa: E731
    v = r["vsan"]
    head = f"""
<section{f' class="{cls}"' if cls else ''} id="vsan"><div class="wrap">
  <div class="eyebrow">If storage moves to vSAN</div>
  <h2>vSAN Host Storage</h2>
  <p class="section-lede">{e(vsan_lede(r, v))}</p>"""
    if v.get("error"):
        return head + "</div></section>"
    best = v["chosen"]
    _, breakdown = vsan_tables_html(r, v)
    cards = f"""
  <div class="cards">
    <div class="card navy"><div class="label">Cluster</div><div class="big">{v['hosts']} <small>hosts</small></div>
      <p>{e(v['policy'])}, vSAN ESA all flash.</p></div>
    <div class="card"><div class="label">Effective needed</div><div class="big">{v['target_fut']:,.2f} <small>TB</small></div>
      <p>Future need. Day 1: {v['target_now']:,.2f} TB.</p></div>
    <div class="card green"><div class="label">Capacity drives per host</div><div class="big">{f"{best['fit']['drives']} x {best['size']:g} <small>TB</small>" if best else 'None'}</div>
      <p>{f"{best['fit']['effective']:,.2f} TB effective from {best['fit']['raw']:,.2f} TB raw." if best else 'Add hosts or drives.'}</p></div>
  </div>"""
    return (head + cards + f"""
  {f'<h3>Capacity breakdown</h3><div class="table-card scroll">{breakdown}</div>' if breakdown else ''}
  <div class="note"><strong>Basis.</strong> {e(VSAN_NOTE)} Compression {v['compression']:.2f}x; up to {v['max_drives']} drives per host.</div>
</div></section>""")


# ---------- raw device mappings (3.13) ----------
LOCAL_LUN_WORDS = ("local", "usb", "boot", "cd-rom", "cdrom")


def rdm_frame(disk, names):
    """RDM disks of the VMs in scope: [VM, Disk, Size (MB), Mode, Sharing, LUN ID, Host, Cluster]."""
    if disk is None or "Raw" not in disk.columns:
        return None
    d = disk[is_true(disk["Raw"])]
    if "VM" in d.columns:
        d = d[d["VM"].astype(str).isin(names)]
    txt = lambda c: d[c].astype(str).replace({"nan": "", "None": ""}) if c in d.columns else pd.Series("", index=d.index)  # noqa: E731
    mode = txt("Raw Comp. Mode").str.replace("Mode", "", case=False).str.strip().str.capitalize()
    share = txt("Sharing mode")
    bus = txt("Shared Bus")
    sharing = share.where(~share.str.lower().isin(["", "sharingnone", "nosharing"]), "")
    sharing = sharing.where(sharing != "", bus.where(~bus.str.lower().isin(["", "nosharing", "none"]), ""))
    nice = {"sharingmultiwriter": "Multi-writer", "physicalsharing": "SCSI bus (physical)",
            "virtualsharing": "SCSI bus (virtual)"}
    sharing = sharing.map(lambda v: nice.get(v.lower(), v))
    size = unit_mb(d, "Capacity")
    return pd.DataFrame({
        "VM": txt("VM"), "Disk": txt("Label").where(txt("Label") != "", txt("Disk")),
        "Size (MB)": size if size is not None else 0.0, "Mode": mode.replace("", "Unknown"),
        "Sharing": sharing.replace("", "None"), "LUN ID": txt("Raw LUN ID"), "Host": txt("Host"),
        "Cluster": txt("Cluster")}).reset_index(drop=True)


def host_lun_frame(mp, clusters, rdms):
    """LUNs presented to hosts that aren't datastores (RVTools vMultiPath): RDM candidates or unused LUNs."""
    if mp is None or mp.empty or "Datastore" not in mp.columns:
        return None
    m = filter_clusters(mp, clusters)
    ds = m["Datastore"].astype(str).str.strip()
    m = m[m["Datastore"].isna() | ds.isin(["", "nan", "None"])]
    key = first_col(m, ["Disk", "UUID", "Display name"])
    if not key or m.empty:
        return pd.DataFrame(columns=["LUN", "Name", "Hosts", "Used as RDM by"])
    name = m["Display name"].astype(str) if "Display name" in m.columns else m[key].astype(str)
    m = m[~name.str.lower().str.contains("|".join(LOCAL_LUN_WORDS), na=False)]
    if m.empty:
        return pd.DataFrame(columns=["LUN", "Name", "Hosts", "Used as RDM by"])
    host = m["Host"].astype(str) if "Host" in m.columns else pd.Series("", index=m.index)
    ids = rdms["LUN ID"].astype(str).str.lower().tolist() if rdms is not None and len(rdms) else []
    vms = rdms["VM"].tolist() if rdms is not None and len(rdms) else []
    rows = []
    for lun, grp in m.groupby(m[key].astype(str)):
        hexid = re.sub(r"^(naa|eui|t10|mpx)\.", "", lun.lower())
        users = sorted({vm for vm, i in zip(vms, ids) if hexid and len(hexid) >= 8 and hexid in i})
        rows.append({"LUN": lun, "Name": str(name.loc[grp.index[0]]), "Hosts": int(host.loc[grp.index].nunique()),
                     "Used as RDM by": ", ".join(users) if users else "Not matched to an RDM in scope"})
    return pd.DataFrame(rows)


def rdm_summary(s):
    rd = s.get("rdms")
    n = len(rd) if rd is not None else 0
    phys = int((rd["Mode"].str.lower() == "physical").sum()) if n else 0
    shared = int((rd["Sharing"] != "None").sum()) if n else 0
    luns = s.get("host_luns")
    unmatched = int((luns["Used as RDM by"] == "Not matched to an RDM in scope").sum()) if luns is not None and len(luns) else 0
    return {"count": n, "vms": int(rd["VM"].nunique()) if n else 0, "physical": phys, "virtual": n - phys if n else 0,
            "shared": shared, "mb": float(rd["Size (MB)"].sum()) if n else 0.0,
            "luns": len(luns) if luns is not None else None, "luns_unmatched": unmatched}


def rdm_flag_text(r):
    x = r["rdm"]
    if not x["count"] and not x["luns_unmatched"]:
        return None
    parts = []
    if x["count"]:
        parts.append(f"{x['count']} RDM disk{'s' if x['count'] != 1 else ''} on {x['vms']} VM{'s' if x['vms'] != 1 else ''} "
                     f"({fmt_cap(x['mb'])}; {x['physical']} physical mode, {x['virtual']} virtual"
                     + (f"; {x['shared']} shared, e.g. Windows failover clusters" if x["shared"] else "") + ")")
    if x["luns_unmatched"]:
        one = x["luns_unmatched"] == 1
        parts.append(f"{x['luns_unmatched']} host LUN{'' if one else 's'} that {'is' if one else 'are'} not "
                     "a datastore or a matched RDM")
    return "; ".join(parts) + ". Plan how these move: physical-mode and shared RDMs can't be Storage vMotioned or converted in place."


def rdm_tables_html(r, table_class="compare sto", limit=200):
    e = lambda v: escape(str(v))  # noqa: E731
    s = r["sto"]
    out = []
    rd = s.get("rdms")
    if rd is not None and len(rd):
        body = "".join(f"<tr><td>{e(x['VM'])}</td><td>{e(x['Disk'])}</td><td>{fmt_cap(x['Size (MB)'])}</td>"
                       f"<td>{e(x['Mode'])}</td><td>{e(x['Sharing'])}</td><td>{e(x['Host'])}</td></tr>"
                       for _, x in rd.head(limit).iterrows())
        more = f"<tr><td colspan='6'>+ {len(rd) - limit:,} more in the app's Raw Data</td></tr>" if len(rd) > limit else ""
        out.append(f"<table class='{table_class}'><tr><th>VM</th><th>Disk</th><th>Size</th><th>Compatibility</th>"
                   f"<th>Sharing</th><th>Host</th></tr>{body}{more}</table>")
    luns = s.get("host_luns")
    if luns is not None and len(luns):
        body = "".join(f"<tr><td>{e(x['Name'])}</td><td>{e(x['LUN'])}</td><td>{x['Hosts']}</td><td>{e(x['Used as RDM by'])}</td></tr>"
                       for _, x in luns.head(limit).iterrows())
        out.append(f"<table class='{table_class}'><tr><th>Host LUN (not a datastore)</th><th>Device</th><th>Hosts</th>"
                   f"<th>Used as RDM by</th></tr>{body}</table>")
    return out


def rdm_section_html(r, cls=""):
    e = lambda v: escape(str(v))  # noqa: E731
    x = r["rdm"]
    tables = rdm_tables_html(r)
    lede = rdm_flag_text(r) or ("No raw device mappings were found for the VMs in scope"
                                + ("" if x["luns"] is not None else " (host LUN data, RVTools vMultiPath, isn't in the export)")
                                + ".")
    cards = (f"""
  <div class="cards">
    <div class="card navy"><div class="label">RDM disks</div><div class="big">{x['count']:,}</div><p>On {x['vms']:,} VM{'' if x['vms'] == 1 else 's'} · {e(fmt_cap(x['mb']))}.</p></div>
    <div class="card"><div class="label">Physical / virtual mode</div><div class="big">{x['physical']:,} <small>/ {x['virtual']:,}</small></div>
      <p>{x['shared']:,} shared (multi-writer or SCSI bus sharing).</p></div>
    <div class="card"><div class="label">Host LUNs not datastores</div><div class="big">{x['luns'] if x['luns'] is not None else 'n/a'}</div>
      <p>{x['luns_unmatched']:,} not matched to an RDM in scope.</p></div>
  </div>""")
    body = "".join(f'<div class="table-card scroll" style="margin-top:20px">{t}</div>' for t in tables)
    return f"""
<section{f' class="{cls}"' if cls else ''} id="rdm"><div class="wrap">
  <div class="eyebrow">Migration check</div>
  <h2>Raw Device Mappings</h2>
  <p class="section-lede">{e(lede)}</p>{cards}{body}
  <div class="note"><strong>Why it matters.</strong> {e(RDM_NOTE)}</div>
</div></section>"""


RDM_NOTE = ("RDMs map a SAN LUN straight to a VM. Virtual-mode RDMs can be converted to VMDKs during Storage vMotion; "
            "physical-mode and shared RDMs (Windows failover clusters, Oracle RAC) stay on the array and must be re-zoned "
            "and re-presented to the new hosts, or migrated inside the guest. Host LUNs that aren't datastores may be RDMs "
            "of VMs outside this scope, or unused. Sources: RVTools vDisk (RDMs) and vMultiPath (host LUNs).")


def cap_big(mb):
    """'1.23 <small>TB</small>' for report cards."""
    num_, unit = fmt_cap(mb).split(" ")
    return f"{num_} <small>{unit}</small>"


def storage_lede(r):
    s = r["sto"]
    t = s["totals"]["inuse"]
    snap = (f" Snapshots hold {fmt_cap(s['snap_mb'])} across {s['snap_cnt']:,} snapshots"
            + (f", {fmt_cap(s['snap_old_mb'])} of it older than {s['old_days']} days" if s["snap_old_cnt"] else "") + "."
            if s["has_snap"] and s["snap_cnt"] else (" No snapshots were found." if s["has_snap"] else
                                                     " Snapshot data isn't in this export."))
    parts = ["VMs"] + (["templates"] if s["tmpl_cnt"] else []) + (["RDM disks"] if s["rdm_cnt"] else []) + \
        (["other datastore files"] if s["other"] else [])
    what = ", ".join(parts[:-1]) + " and " + parts[-1] if len(parts) > 1 else parts[0]
    return (f"{what[0].upper() + what[1:]} occupy {fmt_cap(t['with'])} today "
            f"({fmt_cap(s['totals']['alloc']['with'])} allocated).{snap} Planning for {r['years']} years of growth with "
            f"{s['free_pct']:.0f}% kept free needs about {fmt_cap(s['required']['inuse']['with'][1])} of effective capacity.")


def storage_section_html(r, cls=""):
    e = lambda v: escape(str(v))  # noqa: E731
    s = r["sto"]
    consumed, required = storage_tables_html(r, table_class="compare sto")
    facts = "".join(f"<tr><td>{e(k)}</td><td>{e(v)}</td></tr>" for k, v in storage_facts(r))
    notes = "".join(f" {e(n)}." for n in s["notes"])
    old = (f"<span class='pill warn'>{s['snap_old_cnt']:,} older than {s['old_days']} days</span>" if s["snap_old_cnt"]
           else "<span class='pill ok'>None older than {} days</span>".format(s["old_days"]) if s["snap_cnt"] else "")
    snap_p = (f"{s['snap_cnt']:,} snapshot{'' if s['snap_cnt'] == 1 else 's'} on {s['snap_vms']:,} "
              f"VM{'' if s['snap_vms'] == 1 else 's'}." if s["has_snap"] else "Not in this export.")
    req = s["required"]["inuse"]
    snaps = s["has_snap"] and s["snap_cnt"] > 0
    return f"""
<section{f' class="{cls}"' if cls else ''} id="storage"><div class="wrap">
  <div class="eyebrow">What the data occupies</div>
  <h2>Storage Consumption</h2>
  <p class="section-lede">{e(storage_lede(r))}</p>
  <div class="cards">
    <div class="card navy"><div class="label">In use today</div><div class="big">{cap_big(s['totals']['inuse']['with'])}</div>
      <p>{e(f"Without snapshots: {fmt_cap(s['totals']['inuse']['without'])}." if snaps else "What the VMs and templates occupy on the datastores.")}</p></div>
    {f'<div class="card"><div class="label">Snapshots</div><div class="big">{cap_big(s["snap_mb"])}</div><p>{e(snap_p)}</p>{old}</div>' if snaps else
     f'<div class="card"><div class="label">Allocated</div><div class="big">{cap_big(s["totals"]["alloc"]["with"])}</div><p>Full provisioned size of every disk.{" Snapshot data isn&#39;t in this export." if not s["has_snap"] else ""}</p></div>'}
    <div class="card green"><div class="label">Effective capacity · {r['years']} years</div><div class="big">{cap_big(req['with'][1])}</div>
      <p>{e(f"In use, {r['growth'] * 100:.0f}% growth a year, {s['free_pct']:.0f}% free. Day 1: {fmt_cap(req['with'][0])}." + (f" Without snapshots: {fmt_cap(req['without'][1])}." if snaps else ""))}</p></div>
  </div>
  <div class="table-card scroll">{consumed}</div>
  <h3>Effective capacity required</h3>
  <div class="table-card scroll">{required}</div>
  <div class="card" style="margin-top:20px"><table class="kv">{facts}</table></div>
  <div class="note"><strong>Basis.</strong> {e(STORAGE_NOTE)}{notes}</div>
</div></section>"""


STORAGE_NOTE = ("Effective capacity = consumed x growth / (1 - free-space target): the usable capacity to plan for "
                "before any array data reduction, and without vSAN or RAID protection overhead. In use is what VMs "
                "occupy on the datastores today (thin disks, swap, logs and snapshots); allocated is the full "
                "provisioned size of every disk. Units are decimal (1 TB = 1,000,000 MB).")


# =============================================================================
# 4. PARSERS
# Each parser returns the same "inventory" dict:
#   vms   : DataFrame[name, vcpu, ram_gb]          (in scope)
#   hosts : DataFrame[sockets, cores, ram_gb]      (cores = total per host)
#   storage, vm_prov_tb, vm_used_tb, guest_used_tb, rdm_cnt, rdm_tb
#   has_perf, perf_ghz, perf_estimated, lic_edition, esxi_version
#   raw_vms, raw_hosts (original rows in scope, for the Raw Data tab)
# =============================================================================
BOOT_DS_PATTERN = r"local|boot|service-datastore|scratch|osdata|^datastore\s*\d*(?:\s*\(\d+\))?$"


def drop_boot_datastores(dsdf, vinfo):
    """Remove ESXi boot / local datastores (single-host or boot-style names such as 'datastore1',
    '*-service-datastore1', 'local'), unless a VM or template lives on them (stand-alone hosts keep theirs).
    Returns (datastores kept, number removed)."""
    if "Name" not in dsdf.columns:
        return dsdf, 0
    names = dsdf["Name"].astype(str)
    bootish = names.str.strip().str.lower().str.contains(BOOT_DS_PATTERN, regex=True, na=False)
    if "# Hosts" in dsdf.columns:
        bootish |= num(dsdf["# Hosts"]).fillna(2) <= 1
    homes = set()
    if vinfo is not None and "Path" in vinfo.columns:
        homes = set(vinfo["Path"].astype(str).str.extract(r"^\[([^\]]+)\]")[0].dropna())
    drop = bootish & ~names.isin(homes)
    return dsdf[~drop], int(drop.sum())


def parse_rvtools(sheets, clusters, power=POWER_ALL):
    info_scope = filter_clusters(sheets["vInfo"], clusters)
    info = drop_templates(info_scope)
    vm_on, vm_off = power_counts(info, "Powerstate")
    power = power if power in POWER_MODES else (POWER_ALL if power else POWER_ON)
    info = filter_power(info, "Powerstate", power)

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
    ds = {"cap_mb": 0.0, "used_mb": 0.0, "alloc_mb": 0.0, "mapped": False, "vsan": False}
    dsdf = sheets.get("vDatastore")
    if dsdf is not None and not dsdf.empty:
        d, boot = drop_boot_datastores(dsdf, sheets["vInfo"])
        note = "Shared datastores"
        if clusters is not None:
            ccol = first_col(d, ["Cluster name", "Cluster"])
            if ccol:
                wanted = set(clusters)
                d = d[d[ccol].astype(str).apply(
                    lambda s: any(c.strip() in wanted for c in s.split(",")))]
                note = "Datastores of the selected clusters"
            else:
                note = "All datastores (not mapped to clusters in this export)"
        cap, used = rvtools_tb(d, "Capacity"), rvtools_tb(d, "In Use")
        if boot:
            note += ", ESXi boot volumes excluded"
        storage = {"cap": cap, "used": used, "free": cap - used, "note": note}
        ds.update(cap_mb=mb_sum(d, "Capacity"), used_mb=mb_sum(d, "In Use"), alloc_mb=mb_sum(d, "Provisioned"),
                  mapped=clusters is None or note.startswith("Datastores of the selected"), df=d,
                  ccol=first_col(d, ["Cluster name", "Cluster"]) if clusters is not None else None,
                  vsan="Type" in d.columns and d["Type"].astype(str).str.contains("vsan", case=False).any())

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
        "sto": storage_rvtools(sheets, info_scope, info, clusters, power, ds),
        "vm_prov_tb": rvtools_tb(info, "Provisioned"),
        "vm_used_tb": rvtools_tb(info, "In Use"),
        "guest_used_tb": guest_used_tb, "rdm_cnt": rdm_cnt, "rdm_tb": rdm_tb,
        "has_perf": False, "perf_ghz": 0.0, "perf_estimated": False,
        "lic_edition": lic_edition, "esxi_version": esxi_version,
        "vm_on": vm_on, "vm_off": vm_off, "include_off": power != POWER_ON, "power_scope": power,
        "raw_vms": info, "raw_hosts": vhost,
    }


def parse_live_optics(sheets, clusters, power, basis):
    vm_scope = filter_clusters(sheets["VMs"], clusters)
    vm = drop_templates(vm_scope)
    vm_on, vm_off = power_counts(vm, "Power State")
    power = power if power in POWER_MODES else (POWER_ALL if power else POWER_ON)
    vm = filter_power(vm, "Power State", power)

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
        "ghz": cpu_ghz(esx, ["CPU Speed (MHz)", "Speed (MHz)", "CPU Clock (MHz)"], ["CPU Speed (GHz)", "Speed (GHz)", "CPU Clock Speed (GHz)"]),
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
        if "Device Type" in d.columns:          # local boot/service datastores, unless VMs run on them
            local = d["Device Type"].astype(str).str.lower().eq("local")
            vms_on = num_col(d, "VM Count").fillna(0) > 0
            dropped = int((local & ~vms_on).sum())
            d = d[~local | vms_on]
            if dropped:
                notes.append("ESXi boot volumes excluded")
        id_col = first_col(d, LUN_ID_COLS)
        if id_col:
            d = d.drop_duplicates(subset=[id_col])
        else:
            notes.insert(0, "shared LUNs may be counted once per host")
        storage = {
            "cap": col_sum(d, "Capacity (GiB)") * GIB_MB / 1e6,
            "used": col_sum(d, "Used Capacity (GiB)") * GIB_MB / 1e6,
            "free": col_sum(d, "Free Capacity (GiB)") * GIB_MB / 1e6,
            "note": ", ".join(["Shared datastores"] + notes),
        }

    # Performance
    has_perf, perf_ghz, perf_estimated = False, 0.0, False
    perf = sheets.get("ESX Performance")
    if perf is not None and power != POWER_OFF:      # host-level demand can't be split by VM power state
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
        "sto": {**storage_liveoptics(sheets, vm_scope, vm, power, esx),
                "ds_cap": storage["cap"] * 1e6, "ds_used": storage["used"] * 1e6},
        "vm_prov_tb": col_sum(vm, "Virtual Disk Size (MiB)") / MIB_PER_TB,
        "vm_used_tb": col_sum(vm, "Virtual Disk Used (MiB)") / MIB_PER_TB,
        "guest_used_tb": col_sum(vm, "Guest VM Disk Used (MiB)") / MIB_PER_TB,
        "rdm_cnt": 0, "rdm_tb": 0.0,
        "has_perf": has_perf, "perf_ghz": perf_ghz, "perf_estimated": perf_estimated,
        "lic_edition": lic_edition, "esxi_version": esxi_version,
        "vm_on": vm_on, "vm_off": vm_off, "include_off": power != POWER_ON, "power_scope": power,
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


def design_status(r, hosts, vcpu, ram, ghz=0.0):
    """Failover load for a host count against the design rules."""
    active = hosts - r["ha_nodes"]
    usable_ram = max(active, 0) * r["eff_ram"]
    usable_ghz = max(active, 0) * r.get("eff_ghz", 0.0)
    cpu_ratio = vcpu / (active * r["eff_cores"]) if active > 0 and r["eff_cores"] else 0.0
    ram_use = ram / usable_ram if usable_ram else 0.0
    ghz_use = ghz / usable_ghz if r.get("has_perf") and usable_ghz else None
    return {"active": active, "usable_ram": usable_ram, "cpu_ratio": cpu_ratio, "ram_use": ram_use,
            "usable_ghz": usable_ghz, "ghz_use": ghz_use,
            "cpu_ok": active > 0 and cpu_ratio <= r["vcpu_ratio"] + EPS,
            "ram_ok": active > 0 and ram_use <= 1 + EPS,
            "ghz_ok": ghz_use is None or (active > 0 and ghz_use <= 1 + EPS),
            "ha_ok": active > 0, "min_ok": hosts >= r["min_hosts"]}


def over_hv(r, ratio):
    """True when a vCPU:pCPU ratio is above the selected hypervisor's guidance."""
    return ratio is not None and ratio > r["hv_ratio"] + EPS


def hv_label(r):
    return f"{r['hypervisor']} {r['hv_kind']} {r['hv_ratio']:.1f}:1"


def capacity_checks(r, hosts, vcpu, ram, calc_hosts, ghz=0.0):
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
    ghz_use = ghz / (active * r["eff_ghz"]) if r.get("has_perf") and r.get("eff_ghz") else None
    if ghz_use is not None and ghz_use > 1 + EPS:
        out.append((True, f"Measured CPU: workload needs {ghz_use:.0%} of usable GHz during HA failover "
                          f"({ghz:,.0f} GHz vs {active * r['eff_ghz']:,.0f} GHz at {PERF_TARGET_UTIL:.0%} target)."))
    if not out:
        extra = hosts - calc_hosts
        tail = f" ({extra} host{'s' if extra != 1 else ''} above the calculated {calc_hosts})." if extra > 0 else "."
        ghz_txt = f", {ghz_use:.0%} of usable GHz" if ghz_use is not None else ""
        out.append((False, f"{hl} meets the CPU and RAM design limits with N+{ha} HA{tail} "
                           f"Failover load: {cpu_ratio:.1f}:1 vCPU per usable core, {ram_use:.0%} of usable RAM{ghz_txt}."))
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
    r["sto"] = storage_summary(inv.get("sto") or copy.deepcopy(EMPTY_STO), p, (1 + p["growth"]) ** p["years"])
    r["perf_paired"] = inv.get("perf_paired", [])
    hv = p.get("hypervisor") or DEFAULT_HYPERVISOR
    r["hypervisor"] = hv
    r["hv_ratio"], r["hv_kind"], r["hv_note"] = HYPERVISORS.get(hv, HYPERVISORS[DEFAULT_HYPERVISOR])
    # Usable GHz per host for measured (Live Optics) demand: overhead removed, kept at the target utilisation
    eff_ghz = host_cores * p["clock"] * (1 - p["cpu_buffer"] / 100) * PERF_TARGET_UTIL
    r.update(host_cores=host_cores, eff_cores=eff_cores, eff_ram=eff_ram, eff_ghz=eff_ghz,
             tgt_numa_cores=numa_cores, tgt_numa_ram=numa_ram)
    use_ghz = bool(r["has_perf"]) and eff_ghz > 0

    def size(vcpu, ram, ghz=0.0):
        """Hosts by vCPU ratio, measured GHz (Live Optics) and RAM; the largest wins."""
        h_cpu = ceil_div(vcpu / p["vcpu_ratio"], eff_cores)
        h_ram = ceil_div(ram, eff_ram)
        h_ghz = ceil_div(ghz, eff_ghz) if use_ghz else 0
        need = {"CPU": h_cpu, "GHz": h_ghz, "RAM": h_ram}
        raw = max(need.values())
        top = [k for k, v in need.items() if v == raw]
        constraint = top[0] if len(top) == 1 and raw > 0 else "Balanced"
        total = max(raw + ha, min_hosts)
        return h_cpu, h_ram, h_ghz, raw, constraint, total, (raw + ha) < min_hosts

    def ratios(vcpu, hosts):
        all_hosts = vcpu / (hosts * host_cores) if hosts > 0 else 0.0
        failover = vcpu / ((hosts - ha) * host_cores) if hosts - ha > 0 else None
        return all_hosts, failover

    # Current
    r["hosts_cpu"], r["hosts_ram"], r["hosts_ghz"], r["raw_hosts"], r["constraint"], r["hosts_now"], r["min_applied_now"] = \
        size(r["tot_vcpu"], r["tot_ram"], r["perf_ghz"])
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
    r["fut_ghz"] = r["perf_ghz"] * mult
    r["fut_hosts_cpu"], r["fut_hosts_ram"], r["fut_hosts_ghz"], r["fut_raw_hosts"], r["fut_constraint"], r["hosts_fut"], \
        r["min_applied_fut"] = size(r["fut_vcpu"], r["fut_ram"], r["fut_ghz"])
    r["calc_hosts_fut"] = r["hosts_fut"]
    r["override_fut"] = r["override_now"] and bool(p.get("override_future", True))
    if r["override_fut"]:
        r["hosts_fut"] = int(p["override_hosts"])
    r["ratio_fut_all"], r["ratio_fut_failover"] = ratios(r["fut_vcpu"], r["hosts_fut"])

    # Design-rule validation (most useful when a host count is forced)
    r["checks_now"] = capacity_checks(r, r["hosts_now"], r["tot_vcpu"], r["tot_ram"], r["calc_hosts_now"], r["perf_ghz"])
    r["checks_fut"] = capacity_checks(r, r["hosts_fut"], r["fut_vcpu"], r["fut_ram"], r["calc_hosts_fut"], r["fut_ghz"])
    r["design_now"] = design_status(r, r["hosts_now"], r["tot_vcpu"], r["tot_ram"], r["perf_ghz"])
    r["design_fut"] = design_status(r, r["hosts_fut"], r["fut_vcpu"], r["fut_ram"], r["fut_ghz"])
    r["power_scope"] = power_mode(p)
    add_headroom(r)

    # Performance sizing (Live Optics) - GHz demand, but never below the RAM requirement
    r["perf_hosts_ghz"], r["perf_hosts_rec"], r["perf_constraint"] = 0, 0, ""
    if r["has_perf"]:
        r["perf_hosts_ghz"] = r["hosts_ghz"]
        perf_raw = max(r["perf_hosts_ghz"], r["hosts_ram"])
        r["perf_constraint"] = "CPU (GHz)" if r["perf_hosts_ghz"] >= r["hosts_ram"] else "RAM"
        r["perf_hosts_rec"] = max(perf_raw + ha, min_hosts)

    # Licensing
    r["fut_lic_cores"] = r["hosts_fut"] * license_cores_per_host(p["sockets"], p["cores"])
    r["lic_diff"] = r["fut_lic_cores"] - r["cur_lic_cores"]
    r["vsan"] = vsan_estimate(r)
    r["rdm"] = rdm_summary(r["sto"])
    return r


def constraint_name(c):
    return {"CPU": "CPU (vCPU ratio)", "GHz": "CPU (measured GHz)", "RAM": "RAM"}.get(c, "CPU and RAM balanced")


def sizing_logic_lines(r):
    """Plain-text explanation, shared by the UI and the report."""
    lines = [
        f"Workload: {r['tot_vcpu']:,.0f} vCPU and {r['tot_ram']:,.0f} GB vRAM across {r['tot_vms']:,} VMs.",
        f"Effective host after overhead: {r['eff_cores']:.1f} cores (x {r['vcpu_ratio']:.1f}:1 = "
        f"{r['eff_cores'] * r['vcpu_ratio']:,.0f} vCPU) and {r['eff_ram']:,.0f} GB RAM.",
        (f"Measured CPU: {r['perf_ghz']:,.0f} GHz ({r['lo_basis']}) against {r['eff_ghz']:,.0f} usable GHz per host "
         f"({PERF_TARGET_UTIL:.0%} target)." if r["has_perf"] else None),
        f"Hosts needed: CPU (vCPU ratio) {r['hosts_cpu']}"
        + (f", CPU (measured GHz) {r['hosts_ghz']}" if r["has_perf"] else "")
        + f", RAM {r['hosts_ram']} -> {constraint_name(r['constraint'])} bound, {r['raw_hosts']} active hosts.",
        f"Final: {r['raw_hosts']} + {r['ha_nodes']} HA = {r['raw_hosts'] + r['ha_nodes']}"
        + (f", raised to the {r['min_hosts']}-host minimum cluster size." if r["min_applied_now"] else "."),
        f"Hypervisor: {r['hypervisor']} - {r['hv_ratio']:.1f}:1 vCPU:pCPU is the {r['hv_kind']} guidance. {r['hv_note']}",
    ]
    if r.get("perf_paired"):
        lines.insert(2, "Measured CPU comes from paired Live Optics clusters: "
                     + "; ".join(f"{a} ← {b}" for a, b in r["perf_paired"]) + ". VMs, vCPU and RAM come from RVTools.")
    lines = [x for x in lines if x]
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
.compare tr.total td { font-weight: 600; color: var(--navy); background: rgba(18,58,97,.05); }
.compare tr.sub td { font-size: 13px; }
.compare.sto td:not(:first-child) { white-space: nowrap; }
.compare .tag.mini { font-size: 10.5px; padding: 2px 8px; margin: 0 0 0 6px; vertical-align: middle; }
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
    elif r.get("power_scope", POWER_ALL) == POWER_OFF:
        vm_line = f"{r['tot_vms']:,} powered off for DR ({r['vm_on']:,} powered-on excluded)"
    elif r.get("power_scope", POWER_ALL) == POWER_ON or not r.get("include_off", True):
        vm_line = f"{r['tot_vms']:,} powered on ({r['vm_off']:,} powered-off excluded)"
    else:
        vm_line = f"{r['tot_vms']:,} ({r['vm_on']:,} on / {r['vm_off']:,} off)"

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
    sto, st_ = r["storage"], r["sto"]
    storage = []                       # datastore view; VM-level storage is under Storage consumption
    if sto["cap"] > 0:
        storage += [
            ("Capacity", fmt_cap(sto["cap"] * 1e6), False),
            ("Used", f"{fmt_cap(sto['used'] * 1e6)} ({_pct(sto['used'] / sto['cap'])})", False),
            ("Free", fmt_cap(sto["free"] * 1e6), False),
        ]
    else:
        storage.append(("Datastores", "Not in source file", False))
    storage.append(("vSAN", ("Yes" + (" (VxRail)" if st_.get("vsan_vxrail") else "")) if st_.get("vsan_src") else "No", False))
    storage.append(("RDM disks", f"{r['rdm_cnt']:,}" if r["rdm_cnt"] else "None", bool(r["rdm_cnt"])))
    storage.append(("Scope", sto["note"], False))

    groups = [("Host Inventory", hosts), ("VM Workload", vms),
              ("Density & Utilization", density), ("Datastores", storage)]
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
    constraint_txt = "CPU and RAM balanced" if r["constraint"] == "Balanced" else f"{constraint_name(r['constraint'])} bound"
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
        cons = constraint_name(r["constraint"] if now else r["fut_constraint"])
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
         f"{host_cpu_text(r)}. {r['cpu_buffer']}% CPU and {r['ram_buffer']}% RAM reserved for overhead.",
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
    if d.get("ghz_use") is not None:
        panel_cards.insert(2, ("Measured CPU", f"{d['ghz_use']:.0%} of usable GHz",
                               f"{r['perf_ghz']:,.0f} GHz ({e(r['lo_basis'])}) against {d['usable_ghz']:,.0f} GHz usable during "
                               f"HA failover at a {PERF_TARGET_UTIL:.0%} target.",
                               pill(d["ghz_ok"], "Fits" if d["ghz_ok"] else "Short on CPU")))
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

    extra = [rdm_section_html(r)]
    if r["hypervisor"] == "VMware" and r.get("vsan_show"):
        extra.append(vsan_section_html(r))
    if perf_html:
        extra.append(perf_html)
    extra = [x.replace("<section id=", f'<section class="{"alt" if i % 2 == 0 else ""}" id=', 1) if i % 2 == 0 else x
             for i, x in enumerate(extra)]
    extra_html = "\n".join(extra)
    method_cls = "alt" if len(extra) % 2 == 0 else ""
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
  <a class="link" href="#storage">Storage</a>
  <a class="link" href="#rdm">RDMs</a>
  {'<a class="link" href="#vsan">vSAN</a>' if r['hypervisor'] == 'VMware' and r.get('vsan_show') else ''}
  {perf_nav}
  <button class="btn" onclick="window.print()">Print / Save PDF</button>
</div></nav>

<header class="hero" id="top"><div class="wrap">
  <div class="prepared">
    <span class="eyebrow muted">Prepared for</span>{prepared}
    <span class="divider"></span>
    <span class="eyebrow muted">Scope</span><div class="chip">{e(r['scope_label'])}</div>
  </div>
  <div class="eyebrow kicker">{e(r['hypervisor'])} sizing · {e(r['source_type'])} analysis{' · DR (powered-off VMs)' if r.get('power_scope') == POWER_OFF else ''}{(' · ' + e(r['workload_label'])) if r.get('workload_label') else ''} · {now:%b %d, %Y}</div>
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
  {headroom_cards_html(r)}
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
{storage_section_html(r)}
{extra_html}
{method_section_html([r], method_cls)}
<footer><div class="wrap">
  {f'<img src="{ahead_white}" alt="AHEAD">' if ahead_white else '<strong>AHEAD</strong>'}
  <span>Confidential - Proprietary to AHEAD · Prepared for {e(r['customer'])}</span>
  <span>Source: {e(r['file_name'])} ({e(r['source_type'])})</span>
</div></footer>
</body>
</html>"""



# =============================================================================
# 6b. CONSOLIDATED REPORT (Summary tab + one tab per sized cluster)
# =============================================================================
CONSOLIDATED_CSS = """
.tabbar { position: sticky; top: 64px; z-index: 19; background: var(--white); border-bottom: 1px solid var(--surface); box-shadow: 0 4px 14px rgba(16,32,51,.06); }
.tabbar .wrap { display: flex; flex-wrap: wrap; gap: 6px; padding-top: 10px; padding-bottom: 10px; }
.tab { font: 500 14px var(--font); color: var(--navy); background: transparent; border: 1px solid transparent; border-radius: 9999px;
       padding: 8px 16px; cursor: pointer; white-space: nowrap; display: inline-flex; align-items: center; gap: 8px; }
.tab:hover { background: var(--surface); }
.tab[aria-selected="true"] { background: var(--navy); color: var(--white); }
.tab .kind { font-size: 10px; letter-spacing: .08em; text-transform: uppercase; opacity: .7; }
.tdot { width: 7px; height: 7px; border-radius: 50%; background: #FF8A80; display: inline-block; }
.tabpanel { scroll-margin-top: 130px; }
.table-card.scroll { overflow-x: auto; }
.compare tr.total td { font-weight: 600; color: var(--navy); background: rgba(18,58,97,.05); }
.compare a, .card .label a { color: var(--bright-blue); font-weight: 600; text-decoration: none; }
.tag.mini { font-size: 10.5px; padding: 2px 8px; margin: 0 0 0 6px; vertical-align: middle; }
.foot { font-size: 12.5px; color: var(--gray); margin-top: 12px; }
.flags { list-style: none; padding: 0; margin: 6px 0 0; }
.flags li { font-size: 14px; line-height: 1.55; color: var(--gray); padding: 8px 0; border-bottom: 1px solid var(--surface); }
.flags li:last-child { border-bottom: 0; }
.ftype { display: inline-block; font-size: 11px; font-weight: 600; letter-spacing: .06em; text-transform: uppercase;
         color: var(--white); background: var(--navy); border-radius: 9999px; padding: 2px 9px; margin-right: 8px; }
@media print {
  .tabbar { display: none; }
  .tabpanel[hidden] { display: block !important; }
  .tabpanel + .tabpanel { page-break-before: always; }
}
"""

CONSOLIDATED_JS = """<script>
function showTab(id) {
  document.querySelectorAll('.tabpanel').forEach(function (p) { p.hidden = p.id !== id; });
  document.querySelectorAll('.tab').forEach(function (t) { t.setAttribute('aria-selected', t.dataset.tab === id); });
  if (history.replaceState) history.replaceState(null, '', '#' + id);
  window.scrollTo({ top: 0 });
}
document.querySelectorAll('.tab').forEach(function (t) { t.addEventListener('click', function () { showTab(t.dataset.tab); }); });
var start = (location.hash || '#summary').slice(1);
showTab(document.querySelector('.tabpanel[id="' + start + '"]') ? start : 'summary');
</script>"""

RATIO_CELL_NAMES = {(4, 1): "current", (4, 2): "Day 1", (4, 3): "future",
                    (5, 1): "current at failover", (5, 2): "Day 1 at failover", (5, 3): "future at failover"}


def report_flags(r):
    """Issues worth surfacing on the consolidated Summary tab: [(type, text)]."""
    out = []
    for label, checks in (("Day 1", r["checks_now"]), ("Future", r["checks_fut"])):
        out += [("Shortfall", f"{label}: {t}") for bad, t in checks if bad]
    cells = comparison_flags(r)
    if cells:
        where = ", ".join(RATIO_CELL_NAMES[c] for c in sorted(cells))
        out.append(("Above guidance", f"vCPU:pCPU above the {hv_label(r)} ({where})."))
    out += [("NUMA", t) for bad, t in numa_messages(r) if bad]
    if r.get("rdm") and rdm_flag_text(r):
        out.append(("RDM", rdm_flag_text(r)))
    if r["sto"].get("vsan_src"):
        out.append(("vSAN", vsan_detect_text(r)))
    return out


def report_body(single_html, prefix):
    """Hero + sections of a single-cluster report, with ids prefixed so tabs don't collide."""
    body = single_html[single_html.index('<header class="hero"'):single_html.index("<footer>")]
    body = re.sub(r'<section class="[^"]*" id="method">.*?</section>', "", body, flags=re.S)   # shown once on Summary
    return re.sub(r'id="([^"]+)"', lambda m: f'id="{prefix}-{m.group(1)}"', body)


def generate_consolidated_report(items, customer, customer_logo_uri=None):
    """items: [{"label", "kind": "cluster"|"scenario", "r", "html"}]. Clusters feed the totals;
    scenarios (consolidations / All Clusters) get their own tab but are left out of the totals."""
    e = lambda v: escape(str(v))  # noqa: E731
    f0 = lambda x: f"{x:,.0f}"  # noqa: E731
    ratio = lambda v, c: f"{v / c:.1f}:1" if c else "n/a"  # noqa: E731
    pct = lambda a, b: f"{a / b:.0%}" if b else "n/a"  # noqa: E731
    plural = lambda n, w: f"{n:,} {w}{'' if n == 1 else 's'}"  # noqa: E731
    now = datetime.now()
    ahead_white = file_data_uri(WHITE_LOGO_PATH)
    clusters = [i for i in items if i["kind"] == "cluster"] or items[:1]
    scenarios = [i for i in items if i not in clusters]
    for n, it in enumerate(items):
        it["tab"] = f"t{n}"
        it["flags"] = report_flags(it["r"])
    R = [c["r"] for c in clusters]
    lic_host = lambda r: license_cores_per_host(r["sockets"], r["cores"])  # noqa: E731
    cur_h, d1_h, fu_h = (sum(r[k] for r in R) for k in ("cur_host_count", "hosts_now", "hosts_fut"))
    cur_c = sum(r["cur_cores"] for r in R)
    d1_c, fu_c = sum(r["hosts_now"] * r["host_cores"] for r in R), sum(r["hosts_fut"] * r["host_cores"] for r in R)
    cur_m = sum(r["cur_total_ram_gb"] for r in R)
    d1_m, fu_m = sum(r["hosts_now"] * r["ram"] for r in R), sum(r["hosts_fut"] * r["ram"] for r in R)
    vcpu, vram, fvcpu, fvram = (sum(r[k] for r in R) for k in ("tot_vcpu", "tot_ram", "fut_vcpu", "fut_ram"))
    vms = sum(r["tot_vms"] for r in R)
    cur_l, d1_l, fu_l = sum(r["cur_lic_cores"] for r in R), sum(r["hosts_now"] * lic_host(r) for r in R), sum(r["fut_lic_cores"] for r in R)
    S = [r["sto"] for r in R]
    tsum = lambda b, k: sum(x["totals"][b][k] for x in S)  # noqa: E731
    rsum = lambda b, k, i: sum(x["required"][b][k][i] for x in S)  # noqa: E731
    snap_mb, snap_n = sum(x["snap_mb"] for x in S), sum(x["snap_cnt"] for x in S)
    snap_old_mb, snap_old_n = sum(x["snap_old_mb"] for x in S), sum(x["snap_old_cnt"] for x in S)
    hvs = sorted({r["hypervisor"] for r in R})
    flagged = [c for c in clusters if c["flags"]]
    n_flags = sum(len(c["flags"]) for c in clusters)
    nlabel = plural(len(clusters), "cluster")

    def go(it, text=None):
        return f"<a href='#{it['tab']}' onclick=\"showTab('{it['tab']}');return false;\">{e(text or it['label'])}</a>"

    sto_rows = "".join(
        f"<tr><td>{go(c)}</td><td>{fmt_cap(x['totals']['alloc']['with'])}</td><td>{fmt_cap(x['totals']['inuse']['with'])}</td>"
        f"<td>{fmt_cap(x['snap_mb']) if x['has_snap'] else 'n/a'}</td><td>{fmt_cap(x['snap_old_mb']) if x['has_snap'] else 'n/a'}</td>"
        f"<td>{fmt_cap(x['required']['inuse']['with'][0])}</td><td class='hl'>{fmt_cap(x['required']['inuse']['with'][1])}</td>"
        f"<td>{fmt_cap(x['required']['inuse']['without'][1])}</td></tr>"
        for c, x in zip(clusters, S))
    sto_rows += (f"<tr class='total'><td>Total</td><td>{fmt_cap(tsum('alloc', 'with'))}</td><td>{fmt_cap(tsum('inuse', 'with'))}</td>"
                 f"<td>{fmt_cap(snap_mb)}</td><td>{fmt_cap(snap_old_mb)}</td><td>{fmt_cap(rsum('inuse', 'with', 0))}</td>"
                 f"<td class='hl'>{fmt_cap(rsum('inuse', 'with', 1))}</td><td>{fmt_cap(rsum('inuse', 'without', 1))}</td></tr>")
    vs_rows = ""
    for c in clusters:
        v = c["r"].get("vsan") or {}
        if c["r"]["hypervisor"] != "VMware" or not v or not c["r"].get("vsan_show"):
            continue
        b = v.get("chosen")
        vs_rows += (f"<tr><td>{go(c)}</td><td>{v.get('hosts', '-')}</td><td>{e(v.get('policy', v.get('error', '')))}</td>"
                    + (f"<td>{v['target_fut']:,.2f} TB</td><td class='hl'>"
                       + (f"{b['fit']['drives']} x {b['size']:g} TB</td><td>{b['fit']['effective']:,.2f} TB</td>"
                          f"<td>{b['fit']['raw']:,.2f} TB</td>" if b else "None</td><td>-</td><td>-</td>")
                       if not v.get("error") else "<td>-</td><td>-</td><td>-</td><td>-</td>") + "</tr>")
    vsan_html = (f"<h3>vSAN ESA estimate</h3><div class='table-card scroll'><table class='compare sto'><tr><th>Cluster</th>"
                 "<th>Hosts</th><th>Policy</th><th>Effective needed</th><th>Drives per host</th><th>Effective</th>"
                 f"<th>Raw</th></tr>{vs_rows}</table></div>" if vs_rows else "")
    rd_rows = "".join(
        f"<tr><td>{go(c)}</td><td>{c['r']['rdm']['count']:,}</td><td>{c['r']['rdm']['physical']:,} / {c['r']['rdm']['virtual']:,}</td>"
        f"<td>{c['r']['rdm']['shared']:,}</td><td>{fmt_cap(c['r']['rdm']['mb'])}</td>"
        f"<td>{c['r']['rdm']['luns'] if c['r']['rdm']['luns'] is not None else 'n/a'}</td>"
        f"<td>{'Yes' if c['r']['sto'].get('vsan_src') else 'No'}</td></tr>"
        for c in clusters)
    rdm_html = ("<h3>Raw device mappings and vSAN</h3><div class='table-card scroll'><table class='compare sto'><tr>"
                "<th>Cluster</th><th>RDM disks</th><th>Physical / virtual</th><th>Shared</th><th>RDM size</th>"
                f"<th>Host LUNs not datastores</th><th>vSAN today</th></tr>{rd_rows}</table></div>")
    sto_html = ("<table class='compare sto'><tr><th>Cluster</th><th>Allocated</th><th>In use</th><th>Snapshots</th>"
                "<th>Old snapshots</th><th>Required Day 1</th><th>Required future</th><th>Future without snapshots</th></tr>"
                + sto_rows + "</table>")

    # ---- Summary: hero ----
    stats = [(plural(d1_h, "host"), f"Day 1 across {nlabel} (today {f0(cur_h)})"),
             (plural(fu_h, "host"), "Future requirement with growth"),
             (f"{f0(vcpu)} vCPU", f"{f0(vms)} VMs · {f0(vram)} GB vRAM today"),
             (f"{d1_l - cur_l:+,}" if d1_l != cur_l else "±0", f"Net licensed cores, Day 1 ({f0(cur_l)} → {f0(d1_l)})")]
    stats_html = "".join(f"<div class='stat'><div class='num'>{e(a)}</div><div class='cap'>{e(b)}</div></div>" for a, b in stats)
    prepared = (f"<div class='chip'><img src='{customer_logo_uri}' alt='{e(customer)}'></div>"
                if customer_logo_uri else f"<div class='chip'>{e(customer)}</div>")
    scen_txt = (f" {plural(len(scenarios), 'consolidation scenario')} {'is' if len(scenarios) == 1 else 'are'} shown in "
                f"{'its' if len(scenarios) == 1 else 'their'} own tab{'s' if len(scenarios) != 1 else ''} and not added to the totals."
                if scenarios else "")

    totals = [("Clusters", str(len(clusters)), str(len(clusters)), str(len(clusters))),
              ("Hosts", f0(cur_h), f0(d1_h), f0(fu_h)),
              ("Physical cores", f0(cur_c), f0(d1_c), f0(fu_c)),
              ("Host RAM", f"{f0(cur_m)} GB", f"{f0(d1_m)} GB", f"{f0(fu_m)} GB"),
              ("Workload", f"{f0(vcpu)} vCPU | {f0(vram)} GB", f"{f0(vcpu)} vCPU | {f0(vram)} GB",
               f"{f0(fvcpu)} vCPU | {f0(fvram)} GB"),
              ("vCPU:pCPU (all hosts)", ratio(vcpu, cur_c), ratio(vcpu, d1_c), ratio(fvcpu, fu_c)),
              ("vRAM / host RAM", pct(vram, cur_m), pct(vram, d1_m), pct(fvram, fu_m)),
              ("Licensed cores", f0(cur_l), f0(d1_l), f0(fu_l))]
    totals_html = ("<table class='compare'><tr><th></th><th>Current</th><th>Recommended (Day 1)</th><th>Future</th></tr>"
                   + "".join(f"<tr><td>{e(a)}</td><td>{e(b)}</td><td class='hl'>{e(c)}</td><td>{e(d)}</td></tr>"
                             for a, b, c, d in totals) + "</table>")

    def status(it):
        n = len(it["flags"])
        return "<span class='pill ok'>No issues</span>" if not n else f"<span class='pill warn'>{plural(n, 'flag')}</span>"

    def row(it, total=False):
        r = it["r"]
        tag = " <span class='tag mini'>override</span>" if r["override_now"] else ""
        wl = f"<br><small>{e(r['workload_label'])}</small>" if r.get("workload_label") else ""
        return (f"<tr><td>{go(it)}{wl}</td><td>{e(r['hypervisor'])}</td><td>{f0(r['tot_vms'])}</td><td>{f0(r['cur_host_count'])}</td>"
                f"<td class='hl'>{r['hosts_now']}{tag}</td><td>{r['hosts_fut']}</td><td>{fmt_ratio(r['ratio_now_failover'])}</td>"
                f"<td>{e(constraint_name(r['constraint']))}</td><td>{status(it)}</td></tr>")
    head = ("<tr><th>{}</th><th>Hypervisor</th><th>VMs</th><th>Current hosts</th><th>Day 1 hosts</th><th>Future hosts</th>"
            "<th>Day 1 failover ratio</th><th>Constraint</th><th>Status</th></tr>")
    all_clear = "" if n_flags else "<span class='pill ok'>All clear</span>"
    by_html = ("<table class='compare'>" + head.format("Cluster") + "".join(row(c) for c in clusters)
               + f"<tr class='total'><td>Total</td><td>{e(', '.join(hvs))}</td><td>{f0(vms)}</td><td>{f0(cur_h)}</td>"
               f"<td class='hl'>{f0(d1_h)}</td><td>{f0(fu_h)}</td><td>{ratio(vcpu, d1_c)} all hosts</td><td></td>"
               f"<td>{all_clear}</td></tr></table>")
    scen_html = ""
    if scenarios:
        scen_html = ("<h3>Consolidation scenarios</h3><p class='foot' style='margin-top:-6px'>Alternative ways to size the "
                     "same clusters. Not included in the totals above.</p><div class='table-card scroll'><table class='compare'>"
                     + head.format("Scenario") + "".join(row(s) for s in scenarios) + "</table></div>")

    flag_cards = "".join(
        f"<div class='card'><div class='label'>{go(it, it['label'] + ' →')}</div><ul class='flags'>"
        + "".join(f"<li><span class='ftype'>{e(k)}</span>{e(t)}</li>" for k, t in it["flags"]) + "</ul></div>"
        for it in items if it["flags"])
    clean = [c["label"] for c in clusters if not c["flags"]]
    flags_html = ((f"<div class='cards two'>{flag_cards}</div>" if flag_cards else
                   "<div class='note'><strong>All clear.</strong> No shortfalls, ratios above guidance or wide VMs.</div>")
                  + (f"<div class='note'><strong>No issues:</strong> {e(', '.join(clean))}.</div>" if clean and flag_cards else ""))

    def grow_cell(h):
        if h is None:
            return "n/a"
        return f"{h['vms']:,} ({h['limit']})" if h["vms"] >= 0 else f"<span class='over'>{-h['vms']:,} short</span>"
    grow_rows = "".join(
        f"<tr><td>{go(it)}</td><td>{it['r']['avg_vm_vcpu']:.1f} vCPU · {it['r']['avg_vm_ram']:,.1f} GB</td>"
        f"<td class='hl'>{grow_cell(it['r']['headroom_now'])}</td><td>{grow_cell(it['r']['headroom_fut'])}</td>"
        f"<td>{'~' + format(it['r']['vms_per_host'], ',') if it['r']['vms_per_host'] else 'n/a'}</td></tr>" for it in clusters)
    d1_tot = sum(max(c["r"]["headroom_now"]["vms"], 0) for c in clusters if c["r"]["headroom_now"])
    fu_tot = sum(max(c["r"]["headroom_fut"]["vms"], 0) for c in clusters if c["r"]["headroom_fut"])
    grow_html = ("<table class='compare'><tr><th>Cluster</th><th>Average VM</th><th>Day 1 room (limit)</th>"
                 "<th>Future room (limit)</th><th>Each extra host</th></tr>" + grow_rows
                 + f"<tr class='total'><td>Total</td><td>{(vcpu / vms if vms else 0):.1f} vCPU · {(vram / vms if vms else 0):,.1f} GB</td>"
                 f"<td class='hl'>{d1_tot:,} VMs</td><td>{fu_tot:,} VMs</td><td></td></tr></table>")

    summary = f"""
<header class="hero"><div class="wrap">
  <div class="prepared"><span class="eyebrow muted">Prepared for</span>{prepared}
    <span class="divider"></span><span class="eyebrow muted">Scope</span><div class="chip">{e(nlabel)}</div></div>
  <div class="eyebrow kicker">Solution summary · {e(' / '.join(hvs))} · {now:%b %d, %Y}</div>
  <h1>{e(customer)} host sizing: <em>the whole estate</em></h1>
  <p class="lede">{e(nlabel)} run {f0(vms)} VMs on {f0(cur_h)} hosts today. The recommendation is {plural(d1_h, 'host')} on
    Day 1, growing to {f0(fu_h)} with growth. Each cluster has its own tab with the full sizing.{e(scen_txt)}</p>
  <div class="stats">{stats_html}</div>
  <div class="meta"><span class="dot"></span>Generated {now:%b %d, %Y, %I:%M %p} · {e(APP_TITLE)} v{APP_VERSION}</div>
</div></header>
<section><div class="wrap">
  <div class="eyebrow">Solution totals</div><h2>Current vs. Recommended</h2>
  <p class="section-lede">Totals across {e(nlabel)}. Ratios are worked out from the totals, so they show the overall density.</p>
  <div class="table-card">{totals_html}</div>
</div></section>
<section class="alt"><div class="wrap">
  <div class="eyebrow">By cluster</div><h2>Cluster Comparison</h2>
  <p class="section-lede">Select a name to open its full sizing.</p>
  <div class="table-card scroll">{by_html}</div>
  <p class="foot">Override = host count set manually. Day 1 failover ratio = vCPU:pCPU with the HA hosts failed.</p>
  {scen_html}
</div></section>
<section><div class="wrap">
  <div class="eyebrow">Licensing &amp; storage</div><h2>Roll-up</h2>
  <p class="section-lede">Per-core licensing and storage consumption across {e(nlabel)}, today and after the refresh.</p>
  <div class="cards">
    <div class="card navy"><div class="label">Licensed cores today</div><div class="big">{f0(cur_l)}</div><p>Across {plural(cur_h, 'current host')}.</p></div>
    <div class="card"><div class="label">Licensed cores Day 1</div><div class="big">{f0(d1_l)}</div><p>Future: {f0(fu_l)} cores.</p></div>
    <div class="card {'green' if d1_l <= cur_l else ''}"><div class="label">Net change Day 1</div><div class="big">{d1_l - cur_l:+,}</div>
      <span class='pill {'ok' if d1_l <= cur_l else 'info'}'>{'Fewer or same cores' if d1_l <= cur_l else 'More cores to license'}</span></div>
  </div>
  <h3>Storage</h3>
  <div class="cards">
    <div class="card navy"><div class="label">In use today</div><div class="big">{cap_big(tsum('inuse', 'with'))}</div>
      <p>Allocated: {fmt_cap(tsum('alloc', 'with'))}. Without snapshots: {fmt_cap(tsum('inuse', 'without'))}.</p></div>
    <div class="card"><div class="label">Snapshots</div><div class="big">{cap_big(snap_mb)}</div>
      <p>{snap_n:,} snapshots; {snap_old_n:,} older than the age limit ({fmt_cap(snap_old_mb)}).</p></div>
    <div class="card green"><div class="label">Effective capacity required</div><div class="big">{cap_big(rsum('inuse', 'with', 1))}</div>
      <p>Future, in use with snapshots, per-cluster growth and free space. Without snapshots: {fmt_cap(rsum('inuse', 'without', 1))}.</p></div>
  </div>
  <div class="table-card scroll" style="margin-top:20px">{sto_html}</div>
  <div class="note"><strong>Basis.</strong> {e(STORAGE_NOTE)} Each cluster uses its own growth, free-space target and
    snapshot age limit; datastores shared by several clusters are split between them, so the totals don't count them twice.</div>
  {rdm_html}
  {vsan_html}
</div></section>
<section class="alt"><div class="wrap">
  <div class="eyebrow">Risks &amp; flags</div><h2>What needs attention</h2>
  <p class="section-lede">{plural(n_flags, 'item')} across {plural(len(flagged), 'cluster')}: design-check shortfalls, ratios above
    hypervisor guidance and VMs wider than a socket.</p>
  {flags_html}
</div></section>
<section><div class="wrap">
  <div class="eyebrow">Capacity planning</div><h2>Room to grow</h2>
  <p class="section-lede">How many more VMs of each cluster's average size fit, with the HA hosts failed and the design limit,
    before another host is needed.</p>
  <div class="table-card scroll">{grow_html}</div>
</div></section>
{method_section_html([c["r"] for c in items], "alt")}"""

    panels = [f'<div class="tabpanel" id="summary" role="tabpanel">{summary}</div>']
    tabs = ['<button class="tab" data-tab="summary" role="tab">Summary</button>']
    for it in items:
        panels.append(f'<div class="tabpanel" id="{it["tab"]}" role="tabpanel" hidden>{report_body(it["html"], it["tab"])}</div>')
        kind = "<span class='kind'>scenario</span>" if it in scenarios else ""
        dot = "<span class='tdot'></span>" if it["flags"] else ""
        tabs.append(f'<button class="tab" data-tab="{it["tab"]}" role="tab">{e(it["label"])}{kind}{dot}</button>')
    logo_img = f'<img src="{ahead_white}" alt="AHEAD">' if ahead_white else "AHEAD"
    return f"""<!DOCTYPE html>
<html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">
<title>Host Sizing - {e(customer)} - Consolidated | AHEAD</title>
<link rel="preconnect" href="https://fonts.googleapis.com"><link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Poppins:wght@400;500;600&family=PT+Serif:ital@1&display=swap" rel="stylesheet">
<style>{REPORT_CSS}{CONSOLIDATED_CSS}</style></head>
<body>
<nav class="topnav"><div class="wrap">
  <a class="brand" href="#summary" onclick="showTab('summary');return false;">{logo_img}</a>
  <span style="color:rgba(255,255,255,.8);font-size:14px">{e(customer)} · Consolidated sizing</span>
  <button class="btn" onclick="window.print()">Print / Save PDF</button>
</div></nav>
<div class="tabbar"><div class="wrap" role="tablist">{''.join(tabs)}</div></div>
{''.join(panels)}
<footer><div class="wrap">{logo_img}<span>Confidential - Proprietary to AHEAD · Prepared for {e(customer)}</span>
<span>{e(nlabel)}{' · ' + plural(len(scenarios), 'scenario') if scenarios else ''}</span></div></footer>
{CONSOLIDATED_JS}
</body></html>"""



# =============================================================================
# 7. STREAMLIT UI
# =============================================================================
PRODUCT_NAME = "Host Sizer"
MIN_STREAMLIT = (1, 50)

# Update check: reads only the version line of sizing_app.py on GitHub. Nothing about the
# user or their data is sent. Set HOSTSIZER_NO_UPDATE_CHECK=1 to turn it off.
REPO_URL = "https://github.com/adammarcum/Host-sizing-Tool"
UPDATE_SOURCE_URL = "https://raw.githubusercontent.com/adammarcum/Host-sizing-Tool/main/sizing_app.py"
UPDATE_ZIP_URL = REPO_URL + "/archive/refs/heads/main.zip"
UPDATE_NOTES_URL = REPO_URL + "/blob/main/ReadMe.txt"
MAC_INSTALL_CMD = ("curl -fsSL https://raw.githubusercontent.com/adammarcum/Host-sizing-Tool/main/install_mac.sh"
                   " | bash")

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

[browser]
gatherUsageStats = false
"""

APP_CSS = """<style>
.hs-brand { padding: 2px 2px 18px; margin-bottom: 6px; border-bottom: 1px solid rgba(255,255,255,.12); }
.hs-update { background: rgba(0,179,136,.16); border: 1px solid rgba(110,206,178,.45); border-radius: 12px;
             padding: 12px 14px; margin: 4px 0 10px; color: #FFFFFF; font-size: 13px; line-height: 1.45; }
.hs-update .t { font-weight: 600; color: #6ECEB2; font-size: 12px; letter-spacing: .08em; text-transform: uppercase; }
.hs-update .v { margin: 4px 0 10px; color: rgba(255,255,255,.88); }
.hs-update a { display: inline-block; text-decoration: none; font-weight: 500; font-size: 12.5px; border-radius: 9999px; }
.hs-update a.dl { background: #00B388; color: #FFFFFF; padding: 6px 14px; margin-right: 8px; }
.hs-update a.nt { color: #67C3E9; }
.hs-update .mac { margin-top: 8px; font-size: 11.5px; color: rgba(255,255,255,.7); }
.hs-brand img.lockup { width: 100%; max-width: 230px; height: auto; display: block; }
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
.hs-cmp tr.total td { font-weight: 600; color: #123A61; background: rgba(18,58,97,.05); }
.hs-cmp tr.sub td { font-size: 13px; }
.hs-cmp.sto td:not(:first-child), .hs-cmp.sto th:not(:first-child) { white-space: nowrap; }
.hs-cmp .tag.mini { display: inline-block; font-size: 10.5px; font-weight: 600; color: #009FDC; background: rgba(0,159,220,.12);
                    border-radius: 9999px; padding: 2px 8px; margin-left: 6px; vertical-align: middle; }
.hs-tile-name { font-size: 15px; font-weight: 600; color: #123A61; word-break: break-all; }
</style>"""

GREETING_PARTS = [(12, "morning"), (17, "afternoon"), (24, "evening")]

DEFAULTS = {
    "sockets": 2, "cores": 24, "ram": 1024, "clock": 2.5, "lo_basis": "95th Percentile",
    "hypervisor": DEFAULT_HYPERVISOR, "vcpu_ratio": HYPERVISORS[DEFAULT_HYPERVISOR][0], "cpu_buffer": 10, "ram_buffer": 10, "min_hosts": 2, "ha_nodes": 1,
    "override": False, "override_hosts": 4, "override_future": True,
    "include_off": True, "power_scope": POWER_ALL, "growth_pct": 10.0, "years": 3,
    "customer": "My Customer", "scope": [], "cpu_model": CPU_CUSTOM, "sto_free_pct": 20.0, "snap_old_days": 30,
    "vsan_show": False, "vsan_pick": "Closest fit (automatic)", "vsan_compression": 1.0, "vsan_max_drives": 24, "vsan_raid": VSAN_AUTO, "vsan_src_policy": "RAID-1 mirror, FTT=1 (OSA or ESA) · 2x",
    "cmp_view": "Both", "preview_view": "Consolidated", "upload_mode": "Add to active project", "raw_view": "VMs", "raw_cols": "Key columns", "raw_search": "",
}
PARAM_KEYS = ["hypervisor", "cpu_model", "sockets", "cores", "ram", "clock", "lo_basis", "vcpu_ratio", "cpu_buffer", "ram_buffer",
              "min_hosts", "ha_nodes", "override", "override_hosts", "override_future",
              "include_off", "power_scope", "years", "customer", "sto_free_pct", "snap_old_days",
              "vsan_show", "vsan_pick", "vsan_compression", "vsan_max_drives", "vsan_raid", "vsan_src_policy"]

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


def ensure_streamlit_credentials():
    """Create ~/.streamlit/credentials.toml (blank email) so Streamlit's first-run
    'Email:' prompt doesn't stall the launcher window. Never overwrites an existing file."""
    path = os.path.join(os.path.expanduser("~"), ".streamlit", "credentials.toml")
    try:
        if not os.path.exists(path):
            os.makedirs(os.path.dirname(path), exist_ok=True)
            with open(path, "w", encoding="utf-8") as f:
                f.write('[general]\nemail = ""\n')
    except OSError:
        pass


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
    p["power_scope"] = power_mode(p)
    p["include_off"] = p["power_scope"] != POWER_ON
    return p


def customer_logo_uri():
    logo = st.session_state.get("cust_logo")
    if not logo:
        return None
    return f"data:{logo['type']};base64,{base64.b64encode(logo['bytes']).decode()}"


def report_summary(entry):
    """Quick facts for a Home tile."""
    try:
        facts = [source_facts(src) for src in entry_sources(entry)]
    except Exception:  # noqa: BLE001
        return None, "Could not read file"
    if any(f[0] is None for f in facts):
        return None, "Not an RVTools or Live Optics export"
    n = len(facts)
    lead = f"{n} sources · " if n > 1 else ""
    return (type_label(f[0] for f in facts),
            f"{lead}{sum(f[2] for f in facts):,} hosts · {sum(f[3] for f in facts):,} VMs · {sum(f[1] for f in facts)} clusters")


def safe_filename_part(text):
    """Strip characters Windows and macOS don't allow in file names."""
    cleaned = "".join("_" if ch in '\\/:*?"<>|' or ord(ch) < 32 else ch for ch in str(text))
    return " ".join(cleaned.split()).strip(" .") or "Untitled"


def report_file_name(customer, cluster):
    """Default download name: <customer>-<cluster>-AHEAD Sizing.html"""
    return f"{safe_filename_part(customer)}-{safe_filename_part(cluster)}-AHEAD Sizing.html"


# ---------- multiple sources per project ----------
SRC_SEP = " · "          # cluster id = "<source alias> · <cluster>" when a project has 2+ sources
VCENTER_COLS = {"RVTools": ("vInfo", ["VI SDK Server", "vCenter", "VC Name"]),
                "LiveOptics": ("ESX Hosts", ["vCenter", "vCenter Name", "VCenter", "vCenter Server"])}
VM_NAME_COLS = {"RVTools": ("vInfo", "VM"), "LiveOptics": ("VMs", "VM Name")}
HOST_COLS = {"RVTools": ("vHost", ["Host"]), "LiveOptics": ("ESX Hosts", HOST_NAME_COLS)}


def new_source(file_name, data=None, sha=None, stype=None, sheets=None, alias=None, enabled=True):
    src = {"uid": uuid.uuid4().hex[:10], "file_name": file_name, "sha256": sha, "alias": alias, "enabled": enabled}
    if sheets is not None:
        src.update(type=stype, sheets=sheets)
    else:
        src["bytes"] = data
    return src


def entry_sources(entry):
    """Sources of a session report/project (older single-file entries are converted in place)."""
    if "sources" not in entry:
        if entry.get("sheets") is not None:
            src = new_source(entry.get("file_name", ""), stype=entry.get("source"), sheets=entry["sheets"], sha=entry.get("sha256"))
        else:
            src = new_source(entry.get("file_name", ""), data=entry.get("bytes"), sha=entry.get("sha256"))
        entry["sources"] = [src]
    for src in entry["sources"]:
        if not src.get("alias"):
            src["alias"] = default_alias(src, [s.get("alias") for s in entry["sources"] if s is not src])
    return entry["sources"]


def source_data(src):
    """(type, sheets) for one source; type is None if the file isn't a recognised export."""
    if src.get("sheets") is not None:
        if not src.get("_standalone"):         # projects saved before 3.13
            src["sheets"] = fill_standalone(src["type"], src["sheets"])
            src["_standalone"] = True
        return src["type"], src["sheets"]
    stype, sheets = load_workbook(src["bytes"])
    src["type"] = stype
    return stype, sheets


def default_alias(src, taken=()):
    """Short name for a source: the vCenter host name in the export, else a tidied file name."""
    name = ""
    try:
        stype, sheets = source_data(src)
        sheet, cols = VCENTER_COLS.get(stype, (None, []))
        df = sheets.get(sheet) if sheet else None
        for col in (cols if df is not None else []):
            if col in df.columns:
                vals = df[col].dropna().astype(str).str.strip()
                vals = vals[vals.str.len() > 0]
                cand = str(mode_or(vals, ""))
                if cand and " " not in cand:          # a host name or IP, not a product/version string
                    name = cand.split(".")[0] if not cand.replace(".", "").isdigit() else cand
                    break
    except Exception:  # noqa: BLE001
        name = ""
    if not name:
        stem = os.path.splitext(src.get("file_name") or "Source")[0]
        stem = re.sub(r"^(?i:liveoptics|rvtools)[_\- ]*(\d+[_\- ]*)?", "", stem)          # LiveOptics_3483554_
        stem = re.sub(r"[_\- ]*\d{1,4}[_\-]\d{1,2}[_\-]\d{2,4}([_\- ]*\d{1,2}[.:]\d{2}([.:]\d{2})?)?$", "", stem)  # dates/times
        stem = re.sub(r"[_\- ]*(?i:rvtools|liveoptics|export|all)(?=$|[_\- ])", "", stem)
        stem = re.sub(r"[_\- ]*(?i:rvtools|liveoptics|export|all)(?=$|[_\- ])", "", stem).strip("_- ")
        name = (stem or os.path.splitext(src.get("file_name") or "Source")[0])[:28]
    base, n = name, 2
    while name in taken:
        name, n = f"{base} {n}", n + 1
    return name


def source_ids(entry, include_disabled=False, include_paired=False):
    """[(cluster id, source uid, cluster name)] across the project's sources. Live Optics clusters paired with an
    RVTools cluster are hidden (their measured CPU is used by the RVTools cluster) unless include_paired."""
    srcs = entry_sources(entry)
    multi = len(srcs) > 1
    hidden = set() if include_paired else set((entry.get("pairs") or {}).values())
    out = []
    for src in srcs:
        if not include_disabled and not src.get("enabled", True):
            continue
        try:
            stype, sheets = source_data(src)
        except Exception:  # noqa: BLE001
            continue
        if stype is None:
            continue
        for c in cluster_list(stype, sheets):
            cid = f"{src['alias']}{SRC_SEP}{c}" if multi else c
            if cid not in hidden:
                out.append((cid, src["uid"], c))
    return out


def type_label(types):
    types = sorted(set(t for t in types if t))
    return " + ".join(types) if types else "Unknown"


def active_scope(entry, scope):
    """Scope with out-of-scope (inactive) clusters removed. None (all clusters) becomes the list of active
    clusters when any cluster is inactive. A selection made only of inactive clusters is kept as-is, so
    Current Environment can still show it."""
    inactive = set(entry.get("inactive", []))
    if not inactive:
        return scope
    if scope is None:
        return [cid for cid, _, _ in source_ids(entry) if cid not in inactive]
    kept = [c for c in scope if c not in inactive]
    return kept or scope


def parse_entry(entry, scope, p):
    """Inventory for a scope (list of cluster ids, or None for everything) across all enabled sources.
    Out-of-scope clusters are left out. Returns (inventory, source type label, file names)."""
    srcs = [s for s in entry_sources(entry) if s.get("enabled", True)]
    pairs = active_pairs(entry)
    ids = source_ids(entry, include_paired=True)
    everything = scope is None
    if pairs and scope is None:
        scope = [cid for cid, _, _ in source_ids(entry)]      # paired Live Optics clusters aren't counted twice
    scope = active_scope(entry, scope)
    clustered = {uid for _, uid, _ in ids}
    parts = []
    for src in srcs:
        stype, sheets = source_data(src)
        if stype is None:
            continue
        if scope is None:
            local = None
        elif src["uid"] not in clustered:
            if not everything:
                continue
            local = None                      # no Cluster column: only part of "everything"
        else:
            local = [c for cid, uid, c in ids if uid == src["uid"] and cid in scope]
            if not local:
                continue
        inv = (parse_rvtools(sheets, local, p["power_scope"]) if stype == "RVTools"
               else parse_live_optics(sheets, local, p["power_scope"], p["lo_basis"]))
        if stype == "RVTools" and pairs and local:
            add_paired_perf(entry, src, local, pairs, inv, p)
        parts.append((src, stype, inv))
    if not parts:
        raise ValueError("None of the enabled sources has data for this selection.")
    label = type_label(t for _, t, _ in parts)
    files = ", ".join(s["file_name"] for s, _, _ in parts)
    if len(parts) == 1 and len(entry_sources(entry)) == 1:
        return parts[0][2], label, files
    return merge_inventories(parts), label, files


def add_paired_perf(entry, src, local, pairs, inv, p):
    """Give RVTools clusters the measured CPU (GHz) of their paired Live Optics clusters. Only when every
    cluster in the selection is paired, so the measured demand covers the whole workload."""
    rv_ids = {c: cid for cid, uid, c in source_ids(entry, include_paired=True) if uid == src["uid"]}
    lo_ids = {cid: (uid, c) for cid, uid, c in source_ids(entry, include_paired=True)}
    wanted = [pairs.get(rv_ids.get(c)) for c in local]
    if not wanted or any(w is None or w not in lo_ids for w in wanted):
        return
    by_src = {}
    for w in wanted:
        uid, c = lo_ids[w]
        by_src.setdefault(uid, []).append(c)
    ghz, est, ok = 0.0, False, True
    for uid, names in by_src.items():
        lo_src = next(s for s in entry_sources(entry) if s["uid"] == uid)
        lo_type, lo_sheets = source_data(lo_src)
        lo_inv = parse_live_optics(lo_sheets, names, p["power_scope"], p["lo_basis"])
        ok = ok and lo_inv["has_perf"]
        ghz += lo_inv["perf_ghz"]
        est = est or lo_inv["perf_estimated"]
    if ok:
        inv.update(has_perf=True, perf_ghz=ghz, perf_estimated=est,
                   perf_paired=[(rv_ids[c], pairs[rv_ids[c]]) for c in local])


def merge_inventories(parts):
    """Combine parsed inventories from several sources into one."""
    invs = [inv for _, _, inv in parts]

    def cat(key, tag=False):
        frames = []
        for src, _, inv in parts:
            df = inv[key].copy()
            if tag:
                df.insert(0, "Source", src["alias"])
            frames.append(df)
        return pd.concat(frames, ignore_index=True, sort=False)

    def joined(key):
        vals = list(dict.fromkeys(str(i[key]) for i in invs if str(i[key]) not in ("Unknown", "", "None")))
        return " / ".join(vals) if vals else "Unknown"

    on = [i["vm_on"] for i in invs if i["vm_on"] is not None]
    off = [i["vm_off"] for i in invs if i["vm_off"] is not None]
    stor = [i["storage"] for i in invs]
    has_perf = all(i["has_perf"] for i in invs)
    notes = list(dict.fromkeys(f"{s['alias']}: {i['storage']['note']}" for s, _, i in parts))
    return {
        "vms": cat("vms"), "hosts": cat("hosts"),
        "storage": {"cap": sum(x["cap"] for x in stor), "used": sum(x["used"] for x in stor),
                    "free": sum(x["free"] for x in stor), "note": " | ".join(notes)},
        **{k: sum(i[k] for i in invs) for k in ("vm_prov_tb", "vm_used_tb", "guest_used_tb", "rdm_cnt", "rdm_tb")},
        "sto": combine_storage([i["sto"] for i in invs if i.get("sto")]),
        "has_perf": has_perf, "perf_ghz": sum(i["perf_ghz"] for i in invs) if has_perf else 0.0,
        "perf_estimated": any(i["perf_estimated"] for i in invs),
        "perf_paired": [x for i in invs for x in i.get("perf_paired", [])],
        "lic_edition": joined("lic_edition"), "esxi_version": joined("esxi_version"),
        "vm_on": sum(on) if on else None, "vm_off": sum(off) if off else None,
        "include_off": invs[0].get("include_off", True), "power_scope": invs[0].get("power_scope", POWER_ALL),
        "raw_vms": cat("raw_vms", tag=True), "raw_hosts": cat("raw_hosts", tag=True),
    }


def source_facts(src):
    """(type, n clusters, n hosts, n VMs) for a source tile."""
    stype, sheets = source_data(src)
    if stype is None:
        return None, 0, 0, 0
    vm_sheet, host_sheet = ("vInfo", "vHost") if stype == "RVTools" else ("VMs", "ESX Hosts")
    return (stype, len(cluster_list(stype, sheets)), len(sheets.get(host_sheet, [])),
            len(drop_templates(sheets[vm_sheet])))


def source_names(src):
    """(set of VM names, set of host names) for overlap checks."""
    stype, sheets = source_data(src)
    if stype is None:
        return set(), set()
    vs, vc = VM_NAME_COLS[stype]
    hs, hcols = HOST_COLS[stype]
    vdf, hdf = sheets.get(vs), sheets.get(hs)
    vms = set(vdf[vc].dropna().astype(str).str.lower()) if vdf is not None and vc in vdf.columns else set()
    hc = first_col(hdf, hcols) if hdf is not None else None
    hosts = set(hdf[hc].dropna().astype(str).str.lower()) if hc else set()
    return vms, hosts


def source_overlaps(entry):
    """[(alias a, alias b, shared VMs, shared hosts)] between enabled sources that look like the same environment."""
    srcs = [s for s in entry_sources(entry) if s.get("enabled", True)]
    names = {}
    for s in srcs:
        try:
            names[s["uid"]] = source_names(s)
        except Exception:  # noqa: BLE001
            names[s["uid"]] = (set(), set())
    out = []
    for i, a in enumerate(srcs):
        for b in srcs[i + 1:]:
            v = len(names[a["uid"]][0] & names[b["uid"]][0])
            h = len(names[a["uid"]][1] & names[b["uid"]][1])
            if v or h:
                out.append((a["alias"], b["alias"], v, h))
    return out


def migrate_cluster_ids(entry, before, after):
    """Re-key per-scope settings after sources are added, removed or renamed.
    before/after: {(source uid, cluster name): cluster id}. Scopes that lose a cluster are dropped."""
    mapping = {cid: after.get(k) for k, cid in before.items()}
    lists = entry.get("scope_lists", {})
    touched = set(entry.get("touched", []))

    def conv(scope):
        out = []
        for c in scope:
            n = mapping.get(c, c)
            if n is None:
                return None
            out.append(n)
        return out
    new_s, new_l, new_t = {}, {}, []
    for key, vals in entry.get("scope_settings", {}).items():
        scope = lists.get(key) or ([] if key == ALL_CLUSTERS else key.split(" + "))
        ns = conv(scope)
        if ns is None:
            continue
        nk = scope_key(ns)
        new_s[nk], new_l[nk] = vals, sorted(ns)
        if key in touched:
            new_t.append(nk)
    inactive = [mapping.get(c, c) for c in entry.get("inactive", [])]
    pairs = {mapping.get(a, a): mapping.get(b, b) for a, b in (entry.get("pairs") or {}).items()}
    entry["pairs"] = {a: b for a, b in pairs.items() if a and b}
    labels = {mapping.get(c, c): v for c, v in (entry.get("labels") or {}).items()}
    entry["labels"] = {c: v for c, v in labels.items() if c}
    entry.update(scope_settings=new_s, scope_lists=new_l, touched=new_t,
                 inactive=[c for c in inactive if c],
                 last_scope=conv(entry.get("last_scope", [])) or [])
    entry.pop("_bundle", None)
    return conv


def id_map(entry):
    return {(uid, c): cid for cid, uid, c in source_ids(entry, include_disabled=True, include_paired=True)}


def change_sources(entry_name, entry, change):
    """Apply change(entry) to the source list, then re-key settings and the on-screen scope."""
    ss = st.session_state
    before = id_map(entry)
    change(entry)
    entry_sources(entry)
    conv = migrate_cluster_ids(entry, before, id_map(entry))
    if ss.get("active") == entry_name:
        ss.scope = conv(list(ss.get("scope", []))) or []
        ss._scope_for = None


# ---------- RVTools + Live Optics pairing, workload labels ----------
def active_pairs(entry):
    """Confirmed RVTools→Live Optics cluster pairs whose sources are both present and enabled: {rv id: lo id}."""
    ids = {cid: uid for cid, uid, _ in source_ids(entry, include_paired=True)}
    return {rv: lo for rv, lo in (entry.get("pairs") or {}).items() if rv in ids and lo in ids}


def cluster_fingerprint(stype, sheets, cname):
    """Size facts used to judge whether an RVTools and a Live Optics cluster are the same."""
    def in_cluster(df):
        if df is None:
            return pd.DataFrame()
        return df[df["Cluster"].astype(str) == cname] if "Cluster" in df.columns else df.iloc[0:0]
    if stype == "RVTools":
        vi, vh = in_cluster(drop_templates(sheets["vInfo"])), in_cluster(sheets.get("vHost"))
        vm_col, host_col = "VM", first_col(vh, ["Host"]) if not vh.empty else None
        cores = float((num_col(vh, "# CPU") * num_col(vh, "Cores per CPU")).sum()) if not vh.empty else 0.0
        ram = float(num_col(vh, "# Memory").sum()) / 1024 if not vh.empty else 0.0
    else:
        vi, vh = in_cluster(drop_templates(sheets["VMs"])), in_cluster(sheets.get("ESX Hosts"))
        vm_col, host_col = "VM Name", first_col(vh, HOST_NAME_COLS) if not vh.empty else None
        cores = float(num_col(vh, "CPU Cores").sum()) if not vh.empty else 0.0
        ram = float(num_col(vh, "Memory (KiB)").sum()) / 1024 / 1024 if not vh.empty else 0.0
    names = lambda df, col: set(df[col].dropna().astype(str).str.lower()) if col and col in df.columns else set()  # noqa: E731
    return {"hosts": len(vh), "cores": cores, "ram": ram, "vms": len(vi),
            "vm_names": names(vi, vm_col), "host_names": names(vh, host_col)}


def pair_confidence(a_name, fa, b_name, fb):
    """(label, score, detail) for a candidate pair. Labels: High, Possible, Conflict, Low."""
    def close(x, y):
        return 1.0 if x == y == 0 else (min(x, y) / max(x, y) if max(x, y) > 0 else 0.0)
    metric = min(close(fa[k], fb[k]) for k in ("hosts", "cores", "ram", "vms"))
    def jac(x, y):
        return len(x & y) / len(x | y) if x and y else 0.0
    overlap = max(jac(fa["vm_names"], fb["vm_names"]), jac(fa["host_names"], fb["host_names"]))
    norm = lambda s: "".join(ch for ch in s.lower() if ch.isalnum())  # noqa: E731
    same = norm(a_name) == norm(b_name)
    if overlap >= 0.5 or (same and metric >= 0.9):
        label = "High"
    elif same and metric < 0.75 and overlap < 0.1:
        label = "Conflict"          # same (often anonymised) name, but the clusters look different
    elif same or metric >= 0.9:
        label = "Possible"
    else:
        label = "Low"
    detail = (f"Hosts {fa['hosts']} vs {fb['hosts']} · cores {fa['cores']:,.0f} vs {fb['cores']:,.0f} · "
              f"RAM {fa['ram']:,.0f} vs {fb['ram']:,.0f} GB · VMs {fa['vms']:,} vs {fb['vms']:,}"
              + (f" · {overlap:.0%} names shared" if overlap else ""))
    return label, overlap * 2 + metric + (0.5 if same else 0), detail


def pairing_candidates(entry):
    """RVTools and Live Optics clusters available for pairing: ([(id, name, fp)], [(id, name, fp)])."""
    srcs = {s["uid"]: s for s in entry_sources(entry) if s.get("enabled", True)}
    rv, lo = [], []
    for cid, uid, cname in source_ids(entry, include_paired=True):
        src = srcs.get(uid)
        if not src:
            continue
        stype, sheets = source_data(src)
        target = rv if stype == "RVTools" else lo if stype == "LiveOptics" else None
        if target is not None:
            target.append((cid, cname, cluster_fingerprint(stype, sheets, cname)))
    return rv, lo


def pair_suggestions(entry):
    """{rv id: (lo id, label, detail)} best High/Possible match per RVTools cluster, plus conflicts."""
    rv, lo = pairing_candidates(entry)
    out, conflicts, used = {}, [], set()
    scored = []
    for rid, rname, rfp in rv:
        for lid, lname, lfp in lo:
            label, score, detail = pair_confidence(rname, rfp, lname, lfp)
            if label == "Conflict":
                conflicts.append((rid, lid, detail))
            elif label in ("High", "Possible"):
                scored.append((score, rid, lid, label, detail))
    for score, rid, lid, label, detail in sorted(scored, reverse=True):
        if rid not in out and lid not in used:
            out[rid] = (lid, label, detail)
            used.add(lid)
    return out, conflicts


WORKLOAD_HINTS = [
    (("sql", "mssql"), "SQL Server"), (("ora", "oracle"), "Oracle"), (("ctx", "citrix", "xen", "vda"), "Citrix"),
    (("vdi", "horizon", "hzn"), "VDI"), (("dmz",), "DMZ"), (("epic",), "Epic"), (("sap", "hana"), "SAP"),
    (("exch", "exchange"), "Exchange"), (("gpu", "vgpu"), "GPU"), (("dr",), "DR"),
    (("nonprod", "dev", "test", "qa", "uat"), "Non-production"), (("mgmt", "infra", "inf"), "Infrastructure"),
    (("prod",), "Production"),
]


def suggest_label(entry, cid):
    """Workload label suggested from the cluster name, else from the most common hint in its VM names."""
    tokens = lambda s: set(t for t in re.split(r"[^a-z0-9]+", s.lower()) if t)  # noqa: E731
    found = {c: (uid, name) for c, uid, name in source_ids(entry, include_paired=True)}
    if cid not in found:
        return None, ""
    uid, cname = found[cid]
    def hit(toks, keys):
        return any(t == k or (len(k) >= 3 and t.startswith(k)) for t in toks for k in keys)
    toks = tokens(cname)
    for keys, label in WORKLOAD_HINTS:
        if hit(toks, keys):
            return label, "from the cluster name"
    src = next((s for s in entry_sources(entry) if s["uid"] == uid), None)
    if not src:
        return None, ""
    fp = cluster_fingerprint(*source_data(src), cname)
    if not fp["vm_names"]:
        return None, ""
    counts = {}
    for vm in fp["vm_names"]:
        vt = tokens(vm)
        for keys, label in WORKLOAD_HINTS:
            if hit(vt, keys) or any(k in vm for k in keys if len(k) >= 3):
                counts[label] = counts.get(label, 0) + 1
                break
    if counts:
        label, n = max(counts.items(), key=lambda x: x[1])
        if n / len(fp["vm_names"]) >= 0.3:
            return label, f"{n / len(fp['vm_names']):.0%} of VM names"
    return None, ""


def scope_label_text(entry, scope):
    """Workload label(s) for a selection: the cluster's own label, or the distinct labels of its clusters."""
    labels = entry.get("labels") or {}
    vals = list(dict.fromkeys(labels[c] for c in (scope or []) if labels.get(c)))
    return " / ".join(vals)


def scope_params(values):
    """Full parameter dict for a stored scope (fills gaps with defaults)."""
    ss = st.session_state
    p = {k: DEFAULTS[k] for k in SIZING_KEYS}
    p.update({k: v for k, v in (values or {}).items() if k in SIZING_KEYS})
    p["power_scope"] = power_mode(values or {})
    p["include_off"] = p["power_scope"] != POWER_ON
    p["lo_basis"] = p["lo_basis"] or DEFAULTS["lo_basis"]
    p["hypervisor"] = p["hypervisor"] or DEFAULT_HYPERVISOR
    p["growth"] = p["growth_pct"] / 100
    p["customer"] = ss.customer
    return p


def report_items(entry):
    """Every configured scope -> [{"key", "label", "kind", "r", "html", "file_name"}], ordered:
    single clusters (in source/export order), then consolidations, then All Clusters.
    All Clusters is included only if it was edited, or if no single cluster has been sized."""
    store = entry.get("scope_settings", {})
    lists = entry.get("scope_lists", {})
    touched = set(entry.get("touched", []))
    inactive = set(entry.get("inactive", []))
    clusters = [cid for cid, _, _ in source_ids(entry) if cid not in inactive]
    order = {c: i for i, c in enumerate(clusters)}
    keyed = []
    for key, values in store.items():
        scope = lists.get(key) or ([] if key == ALL_CLUSTERS else [c for c in key.split(" + ") if c in order])
        if any(c in inactive for c in scope):
            continue                                  # selections with an out-of-scope cluster are left out
        scope = [c for c in scope if c in order]
        if key != ALL_CLUSTERS and not scope:
            continue
        keyed.append((key, scope, values))
    singles = sorted([k for k in keyed if len(k[1]) == 1], key=lambda k: order[k[1][0]])
    combos = sorted([k for k in keyed if len(k[1]) > 1], key=lambda k: [order[c] for c in k[1]])
    allc = [k for k in keyed if k[0] == ALL_CLUSTERS and (ALL_CLUSTERS in touched or not singles)]
    items = []
    logo = customer_logo_uri()
    for key, scope, values in singles + combos + allc:
        p = scope_params(values)
        inv, stype, files = parse_entry(entry, scope or None, p)
        r = compute_sizing(inv, p)
        label = " + ".join(scope) if scope else ALL_CLUSTERS
        r.update(scope_label=label, file_name=files, source_type=stype, workload_label=scope_label_text(entry, scope))
        estate = not scope and not singles          # All Clusters is the whole sizing when nothing else is sized
        items.append({"key": key, "label": label, "kind": "cluster" if len(scope) == 1 or estate else "scenario", "r": r,
                      "html": generate_html_report(r, logo), "file_name": report_file_name(p["customer"], label)})
    return items


def report_bundle(entry_name, entry):
    """Consolidated report + every per-scope report, cached until any setting or source changes."""
    ss = st.session_state
    logo = ss.get("cust_logo")
    sig = json.dumps({"s": entry.get("scope_settings", {}), "l": entry.get("scope_lists", {}),
                      "i": sorted(entry.get("inactive", [])),
                      "pr": entry.get("pairs") or {}, "lb": entry.get("labels") or {},
                      "t": sorted(entry.get("touched", [])), "c": ss.customer,
                      "src": [(x["uid"], x["alias"], x.get("enabled", True)) for x in entry_sources(entry)],
                      "logo": hashlib.sha256(logo["bytes"]).hexdigest() if logo else None,
                      "e": entry_name, "v": APP_VERSION}, sort_keys=True, default=str)
    cached = entry.get("_bundle")
    if cached and cached[0] == sig:
        return cached[1]
    items = report_items(entry)
    bundle = {"items": items,
              "html": generate_consolidated_report(items, ss.customer, customer_logo_uri()) if items else None,
              "file_name": report_file_name(ss.customer, "Consolidated")}
    entry["_bundle"] = (sig, bundle)
    return bundle


def build_context():
    """Parse the active project's sources and run the sizing with the current parameters."""
    ss = st.session_state
    name = ss.get("active")
    if not name or name not in ss.reports:
        return None
    entry = ss.reports[name]
    try:
        srcs = entry_sources(entry)
        bad = [s["file_name"] for s in srcs if source_data(s)[0] is None]
    except Exception as exc:  # noqa: BLE001
        return {"error": f"Could not read a file in '{name}'. Is it a valid .xlsx export?", "exc": exc, "name": name}
    if bad:
        return {"error": f"{', '.join(bad)} is not an RVTools export (vInfo tab) or a Live Optics export "
                         f"(VMs + ESX Hosts tabs). Remove it from the project on Home.", "exc": None, "name": name}
    if not any(s.get("enabled", True) for s in srcs):
        return {"error": f"All sources in '{name}' are excluded. Include at least one on Home.", "exc": None, "name": name}
    clusters = [cid for cid, _, _ in source_ids(entry)]
    ss.scope = [c for c in ss.scope if c in clusters]
    if len(clusters) == 1:
        ss.scope = [clusters[0]]          # one cluster: select it rather than "All Clusters"
    sync_scope_settings(name, entry)
    scope = ss.scope or None
    scope_label = ALL_CLUSTERS if scope is None else " + ".join(scope)
    p = params_from_state()
    try:
        inv, stype, files = parse_entry(entry, scope, p)
        r = compute_sizing(inv, p)
        r.update(scope_label=scope_label, file_name=files, source_type=stype,
                 workload_label=scope_label_text(entry, scope))
        html = generate_html_report(r, customer_logo_uri())
    except KeyError as exc:
        return {"error": f"A required column or tab is missing from '{name}': {exc}", "exc": exc, "name": name}
    except Exception as exc:  # noqa: BLE001
        return {"error": f"Error processing '{name}': {exc}", "exc": exc, "name": name}
    report_name = report_file_name(p["customer"], scope_label)
    try:
        bundle = report_bundle(name, entry)
    except Exception as exc:  # noqa: BLE001 - the single report still works
        bundle = {"items": [], "html": None, "file_name": "", "error": str(exc)}
    display = entry.get("project_name") or (name if len(srcs) == 1 else
                                            f"{srcs[0]['file_name']} + {len(srcs) - 1} more source{'s' if len(srcs) > 2 else ''}")
    inactive = set(entry.get("inactive", []))
    out_of_scope = bool(ss.scope) and all(c in inactive for c in ss.scope)
    return {"bundle": bundle, "key": name, "name": display, "source": stype, "inactive": inactive,
            "out_of_scope": out_of_scope,
            "sources": srcs, "clusters": clusters, "inv": inv, "r": r,
            "html": html, "file_name": report_name, "error": None}


# ---------- shared UI pieces ----------
def seg_width(options):
    """Fixed pixel width for a segmented control. Streamlit sizes 'content' width before the
    Poppins web font loads, so the wider font can push the last button onto a second line."""
    return int(sum(len(str(o)) * 9.5 + 40 for o in options) + 12)


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
    if hasattr(st, "space"):          # Streamlit >= 1.51
        st.space("medium")
    else:
        st.html("<div style='height:18px'></div>")
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
                      f":blue-badge[:material/description: {ctx['source']}]"
                      + (f" :blue-badge[:material/stacks: {len(ctx['sources'])} sources]" if len(ctx.get('sources', [])) > 1 else ""),
                      f":gray-badge[:material/hub: {r['scope_label']}]"]
            if r.get("workload_label"):
                badges.append(f":blue-badge[:material/label: {r['workload_label']}]")
            if r.get("perf_paired"):
                badges.append(":green-badge[:material/link: RVTools + Live Optics paired]")
            if r["override_now"]:
                badges.append(":orange-badge[:material/lock: Host count override]")
            if r.get("power_scope") == POWER_ON:
                badges.append(":gray-badge[Powered-on VMs only]")
            elif r.get("power_scope") == POWER_OFF:
                badges.append(":orange-badge[:material/shield: Powered-off VMs only (DR)]")
            badges.append(save_status(st.session_state.reports[ctx["key"]]))
            st.markdown(" ".join(badges))
    with right:
        if ctx and not ctx["error"]:
            render_downloads(ctx)
            save_popover()
    st.divider()


def render_downloads(ctx):
    """Primary: consolidated report (all sized clusters). Secondary: any single cluster/scenario report."""
    bundle = ctx.get("bundle") or {}
    items = bundle.get("items") or []
    if not bundle.get("html"):
        st.download_button("Download report", ctx["html"], file_name=ctx["file_name"], mime="text/html",
                           type="primary", icon=":material/download:", width="stretch")
        return
    n = len(items)
    st.download_button("Download consolidated report", bundle["html"], file_name=bundle["file_name"], mime="text/html",
                       type="primary", icon=":material/download:", width="stretch", key="dl_consolidated",
                       help=f"Summary tab plus {n} cluster tab{'s' if n != 1 else ''} in one file.")
    with st.popover("Single reports", icon=":material/description:", width="stretch"):
        st.caption("Download one cluster or consolidation scenario on its own.")
        listed = set()
        for it in items:
            listed.add(it["key"])
            st.download_button(it["label"] + (" · scenario" if it["kind"] == "scenario" else ""), it["html"],
                               file_name=it["file_name"], mime="text/html", key=f"dl_{it['key']}", width="stretch",
                               icon=":material/download:")
        if scope_key(st.session_state.scope) not in listed:
            st.download_button(f"This selection: {ctx['r']['scope_label']}", ctx["html"], file_name=ctx["file_name"],
                               mime="text/html", key="dl_current", width="stretch", icon=":material/download:")


def context_subtitle(ctx):
    r = ctx["r"]
    return (f"{r['customer']} · {ctx['name']} · {r['cur_host_count']:,} hosts · {r['tot_vms']:,} VMs · "
            f"{r['tot_vcpu']:,.0f} vCPU · {r['tot_ram']:,.0f} GB vRAM")


def scope_bar(ctx):
    if not ctx["clusters"]:
        return
    ss = st.session_state
    inactive = ctx.get("inactive", set())
    help_txt = ("None selected = all clusters. Pick several clusters to model a consolidation. "
                "Clusters marked 'out of scope' are left out of every sizing and report."
                + (f" '{STANDALONE}' groups the ESXi hosts in an export that aren't in a cluster, with their VMs."
                   if any(c.endswith(STANDALONE) for c in ctx["clusters"]) else ""))
    fmt = lambda c: f"{c} · out of scope" if c in inactive else c  # noqa: E731
    if len(ctx["clusters"]) <= 20:
        st.pills("Cluster scope", ctx["clusters"], selection_mode="multi", key="scope", help=help_txt, format_func=fmt)
    else:
        st.multiselect("Cluster scope", ctx["clusters"], key="scope", help=help_txt, placeholder="All clusters",
                       format_func=fmt)
    if len(ss.scope) == 1 and len(ctx["clusters"]) > 1:
        cid = ss.scope[0]
        ss.cluster_active = cid not in inactive
        st.toggle(f"Active: include {cid} in the sizing", key="cluster_active", on_change=_set_cluster_active,
                  args=(ctx["key"], cid),
                  help="Turn off to take this cluster out of scope. It won't be sized, can't be configured, "
                       "and is left out of All Clusters, consolidations and every report.")
    if len(ss.scope) == 1:
        render_label_editor(ctx, ss.scope[0])
    if ctx.get("out_of_scope"):
        st.markdown(":gray-badge[:material/block: Out of scope] This cluster is excluded from every sizing and report. "
                    "Turn **Active** on to include it again.")
    elif len(ss.scope) > 1:
        dropped = [c for c in ss.scope if c in inactive]
        n_in = len(ss.scope) - len(dropped)
        st.caption(f":material/merge: Consolidation scope: combining {n_in} cluster{'s' if n_in != 1 else ''}."
                   + (f" Out of scope and left out: {', '.join(dropped)}." if dropped else ""))
    if inactive and len(ctx["clusters"]) > 1:
        st.caption(f":material/block: Out of scope: {', '.join(c for c in ctx['clusters'] if c in inactive)}")
    store = st.session_state.reports[ctx["key"]].get("scope_settings", {})
    current = scope_key(st.session_state.scope)
    if len(ctx["clusters"]) > 1:
        tags = " ".join(f":{'blue' if k == current else 'gray'}-badge[{k}]" for k in sorted(store))
        c1, c2 = st.columns([6, 1], vertical_alignment="center")
        c1.caption(":material/tune: Sizing settings are kept separately for each cluster selection. "
                   f"Configured: {tags}")
        with c2.popover("Manage", icon=":material/delete_sweep:", width="stretch",
                        help="Remove sizings you don't need, e.g. a combination of clusters changed by accident."):
            st.caption("Remove a sizing to delete its settings and drop it from the reports. Selecting the same "
                       "clusters again starts a fresh sizing.")
            for k in sorted(store):
                a, b = st.columns([5, 1], vertical_alignment="center")
                a.markdown(f":{'blue' if k == current else 'gray'}-badge[{k}]"
                           + (" :gray[(on screen)]" if k == current else ""))
                b.button("", key=f"rmsz_{k}", icon=":material/close:", help=f"Remove the {k} sizing",
                         on_click=_remove_sizing, args=(ctx["key"], k))


def _remove_sizing(entry_name, key):
    """Delete one configured selection's settings. If it's the one on screen, move to another sizing."""
    ss = st.session_state
    entry = ss.reports[entry_name]
    entry.get("scope_settings", {}).pop(key, None)
    entry.get("scope_lists", {}).pop(key, None)
    entry["touched"] = [t for t in entry.get("touched", []) if t != key]
    entry.pop("_bundle", None)
    if scope_key(ss.get("scope", [])) == key:
        remaining = list(entry.get("scope_settings", {}))
        nxt = next((k for k in remaining if k != ALL_CLUSTERS), remaining[0] if remaining else None)
        lists = entry.get("scope_lists", {})
        ss.scope = list(lists.get(nxt) or ([] if nxt in (None, ALL_CLUSTERS) else nxt.split(" + ")))
        ss._scope_for = None
    flash("success", f"Removed the {key} sizing.")


def _set_cluster_active(entry_name, cid):
    ss = st.session_state
    entry = ss.reports[entry_name]
    inactive = [c for c in entry.get("inactive", []) if c != cid]
    if not ss.get("cluster_active"):
        if len(set(inactive) | {cid}) >= len(source_ids(entry)):
            flash("error", "At least one cluster has to stay active.")
            return
        inactive.append(cid)
    entry["inactive"] = inactive
    entry.pop("_bundle", None)


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
# Sizing settings are kept per cluster scope; the customer name is per report/project.
SIZING_KEYS = [k for k in PARAM_KEYS if k != "customer"] + ["growth_pct"]


def scope_key(scope):
    return " + ".join(sorted(scope)) if scope else ALL_CLUSTERS


def sizing_snapshot():
    ss = st.session_state
    return {k: ss[k] for k in SIZING_KEYS}


def load_sizing(values):
    for k, v in (values or {}).items():
        if k in SIZING_KEYS:
            st.session_state[k] = v
    if values:
        st.session_state["power_scope"] = power_mode(values)


def sync_scope_settings(entry_name, entry):
    """Keep the sizing widgets in step with the selected cluster scope.
    Entering a scope loads its saved settings; a scope seen for the first time starts from
    the settings currently on screen (copy last used). Every run stores the current values."""
    ss = st.session_state
    key = scope_key(ss.scope)
    store = entry.setdefault("scope_settings", {})
    if ss.get("_scope_for") != (entry_name, key):
        if key in store:
            load_sizing(store[key])
        ss._scope_for = (entry_name, key)
    elif key in store and store[key] != sizing_snapshot() and key not in entry.setdefault("touched", []):
        entry["touched"].append(key)       # the user changed a setting on this selection
    store[key] = sizing_snapshot()
    entry.setdefault("scope_lists", {})[key] = sorted(ss.scope)
    entry["last_scope"] = list(ss.scope)
    entry["customer"] = ss.customer


def project_settings(entry):
    """Everything a project stores about the user's choices (deep copy)."""
    return project_settings_for(st.session_state.customer, entry)


def project_settings_for(customer, entry):
    return copy.deepcopy({"customer": customer, "scope_settings": entry.get("scope_settings", {}),
                          "scope_lists": entry.get("scope_lists", {}), "touched": sorted(entry.get("touched", [])),
                          "last_scope": entry.get("last_scope", []),
                          "inactive": sorted(entry.get("inactive", [])),
                          "pairs": dict(entry.get("pairs") or {}), "labels": dict(entry.get("labels") or {}),
                          "sources": [[x["uid"], x["alias"], x.get("enabled", True)] for x in entry.get("sources", [])]})


def settings_fingerprint(settings):
    """Part of the settings that counts as a change (which scope is on screen does not)."""
    return {"customer": settings.get("customer"), "scope_settings": settings.get("scope_settings"),
            "sources": settings.get("sources"), "inactive": settings.get("inactive", []),
            "pairs": settings.get("pairs", {}), "labels": settings.get("labels", {})}


def _activate_report(name):
    """Make a session report active and restore its last cluster scope and customer."""
    ss = st.session_state
    ss.active = name
    entry = ss.reports.get(name, {})
    ss.scope = list(entry.get("last_scope", []))
    if entry.get("customer"):
        ss.customer = entry["customer"]
    ss._scope_for = None


def flash(kind, msg):
    """Queue a message from a callback; shown at the top of the next page render."""
    st.session_state._flash = (kind, msg)


def default_project_name(entry_name):
    entry = st.session_state.reports.get(entry_name, {})
    if entry.get("project_name"):
        return entry["project_name"]
    srcs = entry_sources(entry) if entry else []
    base = os.path.splitext(srcs[0]["file_name"] if srcs else entry_name)[0]
    more = f" + {len(srcs) - 1} more" if len(srcs) > 1 else ""
    return f"{st.session_state.customer} - {base}{more}"


def _save_project():
    ss = st.session_state
    name = (ss.get("save_name") or "").strip()
    key = ss.get("active")
    if not name or key not in ss.reports:
        flash("error", "Enter a project name before saving.")
        return
    entry = ss.reports[key]
    try:
        srcs = entry_sources(entry)
        reduced = []
        for src in srcs:
            stype, sheets = source_data(src)
            reduced.append({**{k: v for k, v in src.items() if k not in ("bytes", "sheets", "_bundle")},
                            "type": stype, "sheets": reduce_sheets(stype, sheets)})
        # Guard: the reduced data must produce exactly the same sizing as the full exports.
        p, scope = params_from_state(), (ss.scope or None)
        full_r = compute_sizing(parse_entry(entry, scope, p)[0], p)
        small_r = compute_sizing(parse_entry({**entry, "sources": reduced}, scope, p)[0], p)
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
            "settings": project_settings(entry),
            "customer_logo": ({"type": logo["type"], "name": logo["name"],
                               "b64": base64.b64encode(logo["bytes"]).decode()} if logo else None),
            "sources": [{"uid": r["uid"], "alias": r["alias"], "enabled": r.get("enabled", True), "type": r["type"],
                         "file_name": r["file_name"], "sha256": r.get("sha256"), "sheets": frames_to_json(r["sheets"])}
                        for r in reduced],
            "summary": {"customer": ss.customer, "hypervisor": full_r["hypervisor"],
                        "hosts": full_r["cur_host_count"], "vms": full_r["tot_vms"],
                        "hosts_now": full_r["hosts_now"], "scope": " + ".join(ss.scope) or ALL_CLUSTERS,
                        "configured": sorted(entry.get("scope_settings", {}))},
        }
        write_project(path, payload)
        if old and old != path and os.path.exists(old):
            os.remove(old)   # renamed: replace the previous file
    except Exception as exc:  # noqa: BLE001
        flash("error", f"Could not save the project: {exc}")
        return
    entry.update(project_path=path, project_name=name, saved_settings=project_settings(entry),
                 saved_logo=bool(logo), saved_at=now)
    flash("success", f"Saved '{name}' to {path}")


def _open_project(path):
    ss = st.session_state
    try:
        payload = read_project(path)
        if "sources" in payload:                         # v3: several sources per project
            sources = [new_source(x["file_name"], stype=x["type"], sheets=frames_from_json(x["sheets"]),
                                  sha=x.get("sha256"), alias=x.get("alias"), enabled=x.get("enabled", True))
                       for x in payload["sources"]]
            for src, x in zip(sources, payload["sources"]):
                src["uid"] = x.get("uid") or src["uid"]
        else:                                            # v1/v2: one source
            x = payload["source"]
            sources = [new_source(x["file_name"], stype=x["type"], sheets=frames_from_json(x["sheets"]), sha=x.get("sha256"))]
    except Exception as exc:  # noqa: BLE001
        flash("error", f"Could not open {os.path.basename(path)}: {exc}")
        return
    name = payload["name"]
    key = f"{name} (project)"
    settings = payload.get("settings", {})
    if "scope_settings" in settings:                     # v2: per-scope settings
        scope_settings = settings["scope_settings"]
        last_scope = settings.get("last_scope", [])
    else:                                                # v1: one flat set of settings
        last_scope = settings.get("scope", [])
        scope_settings = {scope_key(last_scope): {k: v for k, v in settings.items() if k in SIZING_KEYS}}
    # Fill any setting a project doesn't carry (older versions) with the tool default.
    scope_settings = {k: {**{sk: DEFAULTS[sk] for sk in SIZING_KEYS}, **{sk: sv for sk, sv in v.items() if sk in SIZING_KEYS}}
                      for k, v in scope_settings.items()}
    customer = settings.get("customer") or payload.get("summary", {}).get("customer") or ss.customer
    logo = payload.get("customer_logo")
    ss.cust_logo = ({"type": logo["type"], "name": logo.get("name", "logo"),
                     "bytes": base64.b64decode(logo["b64"])} if logo else None)
    entry = {"sources": sources, "file_name": sources[0]["file_name"],
             "uploaded": datetime.now(), "project_path": path, "project_name": name,
             "scope_settings": copy.deepcopy(scope_settings), "last_scope": list(last_scope), "customer": customer,
             "scope_lists": copy.deepcopy(settings.get("scope_lists", {})), "touched": list(settings.get("touched", [])),
             "inactive": list(settings.get("inactive", [])),
             "pairs": dict(settings.get("pairs", {})), "labels": dict(settings.get("labels", {})),
             "saved_logo": bool(logo), "saved_at": payload.get("saved")}
    entry_sources(entry)
    entry["saved_settings"] = project_settings_for(customer, entry)
    ss.reports[key] = entry
    _activate_report(key)
    load_sizing(scope_settings.get(scope_key(last_scope)))
    ss._scope_for = (key, scope_key(last_scope))
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
    dirty = (settings_fingerprint(entry.get("saved_settings") or {}) != settings_fingerprint(project_settings(entry))
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
                           f"{sm.get('hosts_now', 0)} recommended · {sm.get('scope', '')}  \n"
                           + (f"{len(sm['configured'])} cluster configuration{'s' if len(sm['configured']) != 1 else ''} · "
                              if sm.get("configured") else "") + f"Saved {saved}")
                b1, b2 = st.columns(2)
                b1.button("Open", key=f"popen_{path}", icon=":material/folder_open:", type="primary",
                          width="stretch", on_click=_open_project, args=(path,))
                with b2.popover("Delete", icon=":material/delete:", width="stretch"):
                    st.caption(f"Delete **{meta['name']}** permanently from this computer?")
                    st.button("Delete project", key=f"pdel_{path}", type="primary", on_click=_delete_project,
                              args=(path,))


# ---------- update check ----------
def version_tuple(v):
    return tuple(int(x) for x in re.findall(r"\d+", str(v))[:4])


@st.cache_data(ttl=6 * 3600, show_spinner=False)
def _fetch_latest_version():
    """Newest APP_VERSION on GitHub. Raises on failure, so failures aren't cached."""
    import urllib.request
    req = urllib.request.Request(UPDATE_SOURCE_URL, headers={"Range": "bytes=0-8191",
                                                             "User-Agent": f"HostSizer/{APP_VERSION}"})
    with urllib.request.urlopen(req, timeout=3) as resp:
        head = resp.read(8192).decode("utf-8", "ignore")
    return re.search(r'^APP_VERSION = "([^"]+)"', head, re.M).group(1)


def latest_version():
    """Newest version on GitHub, or None when disabled, offline or blocked (3 s timeout)."""
    if os.environ.get("HOSTSIZER_NO_UPDATE_CHECK"):
        return None
    try:
        return _fetch_latest_version()
    except Exception:  # noqa: BLE001 - offline, proxy, GitHub down: just skip the notice
        return None


def render_update_notice(container):
    latest = latest_version()
    if not latest or version_tuple(latest) <= version_tuple(APP_VERSION):
        return
    container.html(
        f"<div class='hs-update'><div class='t'>Update available</div>"
        f"<div class='v'>Version {escape(latest)} is out. You have {escape(APP_VERSION)}.</div>"
        f"<a class='dl' href='{UPDATE_ZIP_URL}' target='_blank' rel='noopener'>Download update</a>"
        f"<a class='nt' href='{UPDATE_NOTES_URL}' target='_blank' rel='noopener'>What's new</a>"
        + ("<div class='mac'>On a Mac you can also update by running the install command in Terminal "
           "(see the ReadMe).</div>" if sys.platform == "darwin" else "")
        + "</div>")
    if not st.session_state.get("_update_toast"):
        st.session_state._update_toast = True
        st.toast(f"Host Sizer {latest} is available. See About in the sidebar to download it.", icon=":material/system_update:")


# ---------- sidebar ----------
def render_sidebar():
    sb = st.sidebar
    product_logo = file_data_uri(PRODUCT_LOGO_PATH)
    if product_logo:
        sb.html(f"<div class='hs-brand'><img class='lockup' src='{product_logo}' "
                f"alt='{escape(PRODUCT_NAME)} - Powered by AHEAD'></div>")
    else:  # fallback if the logo file is missing
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
                     on_change=lambda: _activate_report(st.session_state.active))
        if has_report:
            sb.caption(f"{CTX['source']} · {CTX['r']['scope_label']}")
    sb.html("<div class='hs-label'>About</div>")
    sb.caption(f"{APP_TITLE} v{APP_VERSION}")
    render_update_notice(sb)


# ---------- pages ----------
def uploader_key():
    return f"uploader_{st.session_state.get('_up_n', 0)}"


def _ingest_uploads():
    """New project from the uploaded file(s), or add them as sources to the active project."""
    ss = st.session_state
    files = ss.get(uploader_key()) or []
    ss._up_n = ss.get("_up_n", 0) + 1          # fresh uploader next run, so files aren't added twice
    if not files:
        return
    incoming = [(f.name, f.getvalue()) for f in files]
    target = ss.get("active") if ss.get("upload_mode") != "Start a new project" else None
    if target and target in ss.reports:
        entry = ss.reports[target]
        have = {x.get("sha256") for x in entry_sources(entry)}
        new = []
        for name, data in incoming:
            sha = hashlib.sha256(data).hexdigest()
            if sha in have:
                continue
            have.add(sha)
            new.append(new_source(name, data=data, sha=sha))
        if not new:
            flash("error", "Those files are already in this project.")
            return
        change_sources(target, entry, lambda e: e["sources"].extend(new))
        flash("success", f"Added {len(new)} source{'s' if len(new) != 1 else ''} to {target}. Add more, or go to "
                         "Current Environment to start sizing.")
        ss._scope_for = None
        return
    seen, srcs = set(), []
    for name, data in incoming:
        sha = hashlib.sha256(data).hexdigest()
        if sha not in seen:
            seen.add(sha)
            srcs.append(new_source(name, data=data, sha=sha))
    key = srcs[0]["file_name"] if len(srcs) == 1 else f"{srcs[0]['file_name']} + {len(srcs) - 1} more"
    base, n = key, 2
    while key in ss.reports:
        key, n = f"{base} ({n})", n + 1
    ss.reports[key] = {"sources": srcs, "file_name": srcs[0]["file_name"], "uploaded": datetime.now(),
                       "scope_settings": {}, "last_scope": []}
    entry_sources(ss.reports[key])
    _activate_report(key)
    ss.upload_mode = "Add to active project"      # further uploads join this project
    flash("success", f"Started a project with {len(srcs)} source{'s' if len(srcs) != 1 else ''}. Add more files, "
                     "or go to Current Environment to start sizing.")


def _rename_source(entry_name, uid):
    ss = st.session_state
    entry = ss.reports[entry_name]
    new = " ".join(str(ss.get(f"alias_{uid}") or "").split())
    others = {x["alias"] for x in entry_sources(entry) if x["uid"] != uid}
    src = next(x for x in entry["sources"] if x["uid"] == uid)
    if not new or new in others or SRC_SEP.strip() in new:
        flash("error", "Source names must be unique, not empty, and can't contain '·'.")
        ss[f"alias_{uid}"] = src["alias"]
        return
    change_sources(entry_name, entry, lambda e: src.update(alias=new))


def _toggle_source(entry_name, uid):
    ss = st.session_state
    entry = ss.reports[entry_name]
    src = next(x for x in entry["sources"] if x["uid"] == uid)
    src["enabled"] = bool(ss.get(f"inc_{uid}"))
    entry.pop("_bundle", None)
    ss._scope_for = None


def _remove_source(entry_name, uid):
    ss = st.session_state
    entry = ss.reports[entry_name]
    change_sources(entry_name, entry, lambda e: e.update(sources=[x for x in e["sources"] if x["uid"] != uid]))


NO_PAIR = "— not paired —"
CONF_BADGE = {"High": ":green-badge[High]", "Possible": ":blue-badge[Possible]",
              "Conflict": ":red-badge[:material/warning: Conflict]", "Low": ":gray-badge[Low]"}


def _fill_pair_suggestions(entry_name):
    ss = st.session_state
    entry = ss.reports[entry_name]
    sugg, _ = pair_suggestions(entry)
    for rid, (lid, label, _) in sugg.items():
        ss[f"pair_{rid}"] = lid


def _save_pairs(entry_name, rv_ids):
    ss = st.session_state
    entry = ss.reports[entry_name]
    chosen = {rid: ss.get(f"pair_{rid}") for rid in rv_ids}
    chosen = {r: l for r, l in chosen.items() if l and l != NO_PAIR}
    if len(set(chosen.values())) != len(chosen):
        flash("error", "Each Live Optics cluster can be paired with only one RVTools cluster.")
        return
    entry["pairs"] = chosen
    entry.pop("_bundle", None)
    ss._scope_for = None
    flash("success", f"Saved {len(chosen)} pair{'s' if len(chosen) != 1 else ''}. Paired clusters use RVTools for VMs, "
                     "vCPU and RAM and Live Optics for measured CPU.")


def render_pairing(entry_name, entry):
    """Let the user pair RVTools clusters with Live Optics clusters (suggestions + confidence, user confirms)."""
    ss = st.session_state
    try:
        rv, lo = pairing_candidates(entry)
    except Exception:  # noqa: BLE001
        return
    if not rv or not lo:
        return
    sugg, conflicts = pair_suggestions(entry)
    fps = {cid: (name, fp) for cid, name, fp in rv + lo}
    saved = entry.get("pairs") or {}
    with st.container(border=True):
        st.markdown("#### Combine RVTools with Live Optics")
        st.caption("Pair each RVTools cluster with the Live Optics capture of the same cluster: VMs, vCPU and RAM come "
                   "from RVTools, measured CPU (GHz) from Live Optics. Names alone can't be trusted (anonymised exports "
                   "reuse names), so each pair shows a confidence based on host, core, RAM and VM counts and shared names. "
                   "Nothing is paired until you save.")
        lo_opts = [NO_PAIR] + [cid for cid, _, _ in lo]
        for rid, rname, rfp in rv:
            key = f"pair_{rid}"
            if key not in ss or ss[key] not in lo_opts:
                ss[key] = saved.get(rid, NO_PAIR)
            c1, c2, c3 = st.columns([2.2, 2.4, 3], vertical_alignment="center")
            c1.markdown(f"**{rid}**")
            c2.selectbox(f"Live Optics cluster for {rid}", lo_opts, key=key, label_visibility="collapsed")
            pick = ss.get(key)
            if pick and pick != NO_PAIR:
                label, _, detail = pair_confidence(rname, rfp, *fps[pick])
                c3.markdown(f"{CONF_BADGE[label]} :gray[{detail}]")
            elif rid in sugg:
                lid, label, _ = sugg[rid]
                c3.markdown(f":gray[Suggested: {lid}] {CONF_BADGE[label]}")
            else:
                c3.markdown(":gray[No likely match found]")
        for rid, lid, detail in conflicts:
            st.markdown(f":red-badge[:material/warning: Name conflict] **{rid}** and **{lid}** share a name but look like "
                        f"different clusters ({detail}). Possibly anonymised names: only pair them if you're sure.")
        b1, b2, _ = st.columns([1.4, 1.2, 3])
        b1.button("Fill in suggestions", icon=":material/auto_fix_high:", on_click=_fill_pair_suggestions,
                  args=(entry_name,), width="stretch", disabled=not sugg,
                  help="Pre-selects High and Possible matches. Review them, then Save pairs.")
        b2.button("Save pairs", type="primary", icon=":material/link:", on_click=_save_pairs,
                  args=(entry_name, [cid for cid, _, _ in rv]), width="stretch")


def _set_label(entry_name, cid, value=None):
    ss = st.session_state
    entry = ss.reports[entry_name]
    text = " ".join(str(value if value is not None else ss.get("cluster_label") or "").split())[:40]
    labels = dict(entry.get("labels") or {})
    if text:
        labels[cid] = text
    else:
        labels.pop(cid, None)
    entry["labels"] = labels
    entry.pop("_bundle", None)
    if value is not None:
        ss.cluster_label = text


def render_label_editor(ctx, cid):
    ss = st.session_state
    entry = ss.reports[ctx["key"]]
    ss.cluster_label = (entry.get("labels") or {}).get(cid, "")
    c1, c2 = st.columns([2, 3], vertical_alignment="bottom")
    c1.text_input("Workload label", key="cluster_label", on_change=_set_label, args=(ctx["key"], cid),
                  placeholder="e.g. SQL Server, Citrix, DMZ", max_chars=40,
                  help="Shown in the page header and the reports. Leave empty for none.")
    sugg, why = suggest_label(entry, cid)
    with c2:
        if sugg and sugg != ss.cluster_label:
            st.button(f"Use suggestion: {sugg}", icon=":material/lightbulb:", on_click=_set_label,
                      args=(ctx["key"], cid, sugg), help=f"Suggested {why}. Check it before using it.", key="label_sugg")


def render_project_sources():
    """Sources card for the active project: rename, include/exclude, remove, overlap warnings."""
    ss = st.session_state
    name = ss.get("active")
    if not name or name not in ss.reports:
        return
    entry = ss.reports[name]
    try:
        srcs = entry_sources(entry)
    except Exception:  # noqa: BLE001
        return
    section_title("Active project", f"Sources in {entry.get('project_name') or name}")
    st.caption("Every cluster from these files is listed under Cluster scope on the sizing pages. With two or more "
               "files, clusters are named \"source · cluster\"; rename a source below. Upload more files above to add them, "
               "then click Configure clusters. Save project keeps the whole collection.")
    with st.container(border=True):
        for src in srcs:
            uid = src["uid"]
            if f"alias_{uid}" not in ss:
                ss[f"alias_{uid}"] = src["alias"]
            if f"inc_{uid}" not in ss:
                ss[f"inc_{uid}"] = src.get("enabled", True)
            try:
                stype, n_cl, n_h, n_vm = source_facts(src)
            except Exception:  # noqa: BLE001
                stype, n_cl, n_h, n_vm = None, 0, 0, 0
            c1, c2, c3, c4 = st.columns([2.2, 3, 1.3, 1], vertical_alignment="center")
            c1.text_input("Source name", key=f"alias_{uid}", on_change=_rename_source, args=(name, uid),
                          label_visibility="collapsed")
            c2.markdown(f":{'blue' if stype else 'red'}-badge[{stype or 'Not recognised'}] "
                        f":gray[{escape(src['file_name'])} · {n_cl} clusters · {n_h:,} hosts · {n_vm:,} VMs]")
            c3.toggle("Include", key=f"inc_{uid}", on_change=_toggle_source, args=(name, uid),
                      help="Excluded sources are kept in the project but left out of every sizing and report.")
            c4.button("Remove", key=f"rmsrc_{uid}", icon=":material/delete:", disabled=len(srcs) < 2,
                      on_click=_remove_source, args=(name, uid), width="stretch")
    if st.button("Configure clusters", type="primary", icon=":material/arrow_forward:", key="go_env"):
        st.switch_page(PAGES["env"])
    render_pairing(name, entry)
    paired_aliases = set()
    ids = {cid: uid for cid, uid, _ in source_ids(entry, include_paired=True)}
    alias_of = {x["uid"]: x["alias"] for x in srcs}
    for rv, lo in active_pairs(entry).items():
        paired_aliases.add(frozenset((alias_of.get(ids.get(rv)), alias_of.get(ids.get(lo)))))
    for a, b, v, h in source_overlaps(entry):
        if frozenset((a, b)) in paired_aliases:
            continue
        what = " and ".join(x for x in (f"{v:,} VMs" if v else "", f"{h:,} hosts" if h else "") if x)
        st.markdown(f":orange-badge[:material/warning: Possible overlap] **{a}** and **{b}** share {what}. They may "
                    "describe the same environment; exclude one to avoid counting it twice.")


def _open_report(name):
    _activate_report(name)
    st.session_state._goto = "env"


def _remove_report(name):
    ss = st.session_state
    ss.reports.pop(name, None)
    if ss.active == name:
        nxt = next(iter(ss.reports), None)
        if nxt:
            _activate_report(nxt)
        else:
            ss.active, ss.scope = None, []


def page_home():
    n = len(st.session_state.reports)
    page_header(greeting(), f"{n} report{'s' if n != 1 else ''} loaded this session" if n
                else "Upload an RVTools or Live Optics export to start a sizing.")
    with st.container(border=True):
        active = st.session_state.get("active")
        has_active = bool(active and active in st.session_state.reports)
        st.markdown("#### Add source files")
        st.caption("RVTools (.xlsx with a vInfo tab) or Dell Live Optics (.xlsx with VMs and ESX Hosts tabs). Add as many "
                   "as you need: all their clusters appear together under Cluster scope. Files stay on this computer.")
        if has_active:
            if st.session_state.get("upload_mode") not in ("Add to active project", "Start a new project"):
                st.session_state.upload_mode = "Add to active project"
            st.segmented_control("Uploads go to", ["Add to active project", "Start a new project"], key="upload_mode",
                                 width=seg_width(["Add to active project", "Start a new project"]),
                                 help=f"'Add to active project' adds files as extra sources to '{active}'. "
                                      "'Start a new project' puts the next upload in a separate project.")
        st.file_uploader("Upload RVTools or Live Optics Excel file(s)", type=["xlsx"], accept_multiple_files=True,
                         key=uploader_key(), on_change=_ingest_uploads, label_visibility="collapsed")

    if n:
        render_project_sources()
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
    render_storage(r)


def render_storage(r):
    s = r["sto"]
    section_title("What the data occupies", "Storage consumption")
    if s.get("vsan_src"):
        st.markdown(":blue-badge[:material/storage: vSAN detected] " + vsan_detect_text(r)
                    + (" Change this under Sizing parameters > Storage." if s.get("vsan_copies_in_vm") else ""))
    st.caption(storage_lede(r))
    t, req = s["totals"], s["required"]["inuse"]
    c = st.columns(4)
    snaps = s["has_snap"] and s["snap_cnt"] > 0
    c[0].metric("In use today", fmt_cap(t["inuse"]["with"]), border=True,
                help=(f"Without snapshots: {fmt_cap(t['inuse']['without'])}" if snaps else
                      "What the VMs and templates occupy on the datastores"))
    c[1].metric("Allocated", fmt_cap(t["alloc"]["with"]), border=True, help="Full provisioned size of every disk")
    c[2].metric("Snapshots", fmt_cap(s["snap_mb"]) if s["has_snap"] else "Not in export", border=True,
                help=(f"{s['snap_cnt']:,} snapshots on {s['snap_vms']:,} VMs" if s["has_snap"] else
                      "This export has no snapshot data"))
    c[3].metric(f"Effective required · {r['years']} yr", fmt_cap(req["with"][1]), border=True,
                help=f"In use, {r['growth'] * 100:.0f}% growth a year, {s['free_pct']:.0f}% free. Day 1: "
                     f"{fmt_cap(req['with'][0])}." + (f" Without snapshots: {fmt_cap(req['without'][1])}." if snaps else ""))
    if s["snap_old_cnt"]:
        st.markdown(f":orange-badge[:material/history: {s['snap_old_cnt']:,} snapshots older than {s['old_days']} days] "
                    f"{fmt_cap(s['snap_old_mb'])}, oldest {s['snap_oldest']:,} days. They may be unintentional; "
                    "compare the *Without old snapshots* column.")
    consumed, required = storage_tables_html(r, table_class="hs-cmp sto", hl_class="hs-hl")
    with st.container(border=True):
        st.html(consumed)
    with st.container(border=True):
        st.html(required)
    with st.container(border=True):
        st.html(kv_html([(k, v, False) for k, v in storage_facts(r)]))
    st.caption(STORAGE_NOTE + "".join(f" {n}." for n in s["notes"])
               + " Free space and the snapshot age limit are under Sizing parameters > Storage.")
    render_rdm(r)


def render_rdm(r):
    x, s = r["rdm"], r["sto"]
    section_title("Migration check", "Raw device mappings")
    text = rdm_flag_text(r)
    if text:
        st.markdown(f":orange-badge[:material/link: RDMs found] {text}")
    else:
        st.caption("No raw device mappings were found for the VMs in scope"
                   + ("" if x["luns"] is not None else " (host LUN data, RVTools vMultiPath, isn't in this export)") + ".")
    c = st.columns(4)
    c[0].metric("RDM disks", f"{x['count']:,}", border=True, help=f"On {x['vms']:,} VMs")
    c[1].metric("Physical / virtual", f"{x['physical']:,} / {x['virtual']:,}", border=True)
    c[2].metric("Shared RDMs", f"{x['shared']:,}", border=True, help="Multi-writer or SCSI bus sharing (e.g. WSFC)")
    c[3].metric("Host LUNs not datastores", x["luns"] if x["luns"] is not None else "n/a", border=True,
                help=f"{x['luns_unmatched']:,} not matched to an RDM in scope")
    rd, luns = s.get("rdms"), s.get("host_luns")
    if rd is not None and len(rd):
        view = rd.assign(**{"Size": rd["Size (MB)"].map(fmt_cap)}).drop(columns=["Size (MB)"])
        st.dataframe(view[["VM", "Disk", "Size", "Mode", "Sharing", "Host", "Cluster", "LUN ID"]], hide_index=True,
                     width="stretch")
    if luns is not None and len(luns):
        st.dataframe(luns.rename(columns={"Name": "Host LUN (not a datastore)", "LUN": "Device"}), hide_index=True,
                     width="stretch")
    st.caption(RDM_NOTE)


def _apply_hv_ratio():
    """Set the design limit to the selected hypervisor's vCPU:pCPU guidance."""
    hv = st.session_state.get("hypervisor") or DEFAULT_HYPERVISOR
    st.session_state.vcpu_ratio = HYPERVISORS[hv][0]


def _ram_label(v):
    try:
        return f"{int(v):,} GB"
    except (TypeError, ValueError):
        return str(v)


def _ram_entered():
    """RAM typed by hand: accept '1000', '1,000' or '1000 GB'; otherwise keep the last good value."""
    ss = st.session_state
    raw = str(ss.get("ram") or "").upper().replace(",", "").replace("GB", "").strip()
    try:
        val = int(float(raw))
    except ValueError:
        val = None
    if val is None or not RAM_MIN <= val <= RAM_MAX:
        ss.ram = int(ss.get("_ram_ok") or DEFAULTS["ram"])
        flash("error", f"RAM per host must be a number of GB between {RAM_MIN:,} and {RAM_MAX:,}.")
        return
    ss.ram = val
    ss._ram_ok = val


def _apply_cpu():
    """Processor picked: copy its cores per socket and base GHz into Target hardware."""
    ss = st.session_state
    c = CPU_BY_LABEL.get(ss.get("cpu_model") or "")
    if not c:
        return
    ss.cores = c["cores"]
    ss.clock = float(c["ghz"])
    if c["max_sockets"] == 1:
        ss.sockets = 1


def _cpu_edited():
    """Cores or GHz typed by hand: drop the processor name once they no longer match it."""
    ss = st.session_state
    if ss.get("cpu_model") != CPU_CUSTOM and selected_cpu(ss) is None:
        ss.cpu_model = CPU_CUSTOM


def render_parameters(ctx):
    with st.container(border=True):
        st.markdown("#### Sizing parameters")
        st.caption(f"These settings apply to **{scope_key(st.session_state.scope)}** only. Each cluster selection "
                   "keeps its own settings; a new selection starts from the last one you used.")
        t0, t1, t2, t3, t4, t5 = st.tabs([":material/layers: Hypervisor", ":material/memory: Target hardware",
                                          ":material/rule: Constraints", ":material/lock: Host count override",
                                          ":material/trending_up: Scope & growth", ":material/storage: Storage"])
        with t0:
            st.segmented_control("Target hypervisor", list(HYPERVISORS), key="hypervisor", on_change=_apply_hv_ratio,
                                 width=seg_width(HYPERVISORS),
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
            ss = st.session_state
            if ss.get("cpu_model") not in CPU_BY_LABEL:
                ss.cpu_model = CPU_CUSTOM
            st.selectbox("Processor", CPU_OPTIONS, key="cpu_model", on_change=_apply_cpu,
                         help="Intel Xeon 6 and 5th Gen Xeon, AMD EPYC 9005 and 9004, sorted by generation and core count. "
                              "Type to search (e.g. 6737P or 9375F). Picking one fills Cores / socket and the base clock; "
                              "Custom keeps your own values.")
            cpu = CPU_BY_LABEL.get(ss.cpu_model)
            c = st.columns(4)
            c[0].number_input("Sockets / host", 1, 4, key="sockets")
            c[1].number_input("Cores / socket", 4, 256, key="cores", on_change=_cpu_edited)
            try:
                ss.ram = int(ss.ram)
            except (TypeError, ValueError):
                ss.ram = int(ss.get("_ram_ok") or DEFAULTS["ram"])
            ss._ram_ok = ss.ram
            ram_opts = sorted(set(RAM_OPTIONS) | {ss.ram})
            c[2].selectbox("RAM / host (GB)", ram_opts, key="ram", format_func=_ram_label, accept_new_options=True,
                           on_change=_ram_entered,
                           help="Common configurations, multiples of 256 GB up to 4 TB. Type another amount "
                                f"({RAM_MIN:,}-{RAM_MAX:,} GB) and press Enter to use it.")
            c[3].number_input("CPU speed (GHz)", 1.0, 5.0, key="clock", format="%.2f", on_change=_cpu_edited,
                              help="Base clock per core. Used for measured CPU (Live Optics) sizing.")
            if int(ss.ram) % 256:
                st.markdown(f":orange-badge[:material/warning: {int(ss.ram):,} GB isn't a multiple of 256 GB] "
                            "Current servers fill every memory channel evenly; check the DIMM layout supports it.")
            if cpu:
                st.caption(f"**{cpu['name']}**: {cpu['cores']} cores per socket, {cpu['ghz']:.2f} GHz base, "
                           f"up to {cpu['max_sockets']} socket{'' if cpu['max_sockets'] == 1 else 's'}"
                           + (f" ({cpu['note']})" if cpu["note"] else "") + ".")
                if ss.sockets > cpu["max_sockets"]:
                    st.markdown(f":orange-badge[:material/warning: {cpu['name']} supports {cpu['max_sockets']} "
                                f"socket{'' if cpu['max_sockets'] == 1 else 's'}] Lower Sockets / host or pick another processor.")
            st.segmented_control(
                "Live Optics basis", ["95th Percentile", "Peak CPU", "Average CPU"], key="lo_basis", width=seg_width(["95th Percentile", "Peak CPU", "Average CPU"]),
                disabled="LiveOptics" not in ctx["source"],
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
            st.segmented_control("VMs to size", POWER_MODES, key="power_scope", width=seg_width(POWER_MODES),
                                 help="All VMs, only powered-on VMs, or only powered-off VMs (use this to size a DR "
                                      "environment). Suspended VMs count as powered off. Live Optics performance "
                                      "sizing isn't available for powered-off only.")
        with t5:
            c = st.columns(2)
            c[0].number_input("Storage free space (%)", 0.0, 90.0, step=5.0, key="sto_free_pct",
                              help="Capacity kept free on new array storage. Effective capacity = consumed x growth / "
                                   "(1 - free space). vSAN uses its reserves instead (except 2-host clusters).")
            c[1].number_input("Old snapshot age (days)", 1, 3650, key="snap_old_days",
                              help="Snapshots older than this are flagged as possibly unintentional and shown in the "
                                   "'Without old snapshots' column.")
            ss = st.session_state
            ss.vsan_src_policy = VSAN_SRC_LEGACY.get(ss.get("vsan_src_policy"), ss.get("vsan_src_policy"))
            if ss.vsan_src_policy not in VSAN_SRC_POLICIES:
                ss.vsan_src_policy = DEFAULTS["vsan_src_policy"]
            if ss.get("vsan_raid") not in [VSAN_AUTO] + list(VSAN_POLICIES):
                ss.vsan_raid = VSAN_AUTO
            st.toggle("Show vSAN options", key="vsan_show",
                      help="Adds a vSAN ESA host storage estimate to the Recommendation page and the reports "
                           "(VMware targets). vSAN found in the source files is always reported.")
            if not ss.vsan_show:
                st.caption("Turn on to estimate vSAN ESA drives per host for this sizing.")
            else:
                c = st.columns(3)
                c[0].selectbox("New cluster protection (vSAN ESA)", [VSAN_AUTO] + list(VSAN_POLICIES), key="vsan_raid",
                               help="Auto picks by host count: 2 hosts RAID-1 mirror, 3-5 hosts RAID-5 2+1, 6+ hosts "
                                    "RAID-6 4+2. RAID-5 4+1 (6+ hosts) is the FTT=1 alternative for larger clusters.")
                c[1].number_input("Compression ratio", 1.0, 3.0, step=0.1, key="vsan_compression",
                                  help="1.0 = no data-reduction credit. Broadcom's sizer often uses 1.2.")
                c[2].number_input("Max drives per host", VSAN_MIN_DRIVES, 40, key="vsan_max_drives",
                                  help="Chassis limit for capacity drives.")
            sto = (ctx.get("r") or {}).get("sto") or {}
            if sto.get("vsan_src") and sto.get("vsan_copies_in_vm"):
                st.selectbox("Current environment: vSAN policy today", list(VSAN_SRC_POLICIES), key="vsan_src_policy",
                             help="RVTools counts a VM's usage on vSAN including its protection copies (e.g. a 100 GB "
                                  "VM on RAID-1 shows 200 GB). Pick today's policy so the new storage is sized on the "
                                  "data itself. 'No adjustment' keeps the usage as reported.")


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
        badges = [f":blue-badge[N+{r['ha_nodes']}]", f":gray-badge[{constraint_name(cons)} constraint]"]
        if tag:
            badges.append(f":orange-badge[:material/lock: {tag}]")
        st.markdown(" ".join(badges))
        st.markdown(f"vCPU:pCPU {ratio_md(r, all_r)} across all hosts · {ratio_md(r, fo_r)} during HA failover  \n"
                    f":gray[Design limit {r['vcpu_ratio']:.1f}:1 · {hv_label(r)}]")
        if over_hv(r, all_r) or over_hv(r, fo_r):
            st.markdown(f":red-badge[:material/warning: Above {hv_label(r)}]")
        render_checks(r["checks_now"] if now else r["checks_fut"])


def render_vsan(r):
    v = r["vsan"]
    section_title("If storage moves to vSAN", "vSAN ESA host storage estimate")
    st.caption(vsan_lede(r, v))
    if v.get("error"):
        return
    best = v["chosen"]
    c = st.columns(4)
    c[0].metric("Hosts · policy", f"{v['hosts']} · {v['policy'].split(' (')[0]}", border=True, help=v["policy"])
    c[1].metric("Effective needed", f"{v['target_fut']:,.2f} TB", border=True,
                help=f"Day 1 need: {v['target_now']:,.2f} TB")
    c[2].metric("In report / host", f"{best['fit']['drives']} x {best['size']:g} TB" if best else "None", border=True)
    c[3].metric("Effective provided", f"{best['fit']['effective']:,.2f} TB" if best else "n/a", border=True,
                help=(f"Raw {best['fit']['raw']:,.2f} TB" if best else None))
    options, breakdown = vsan_tables_html(r, v, table_class="hs-cmp sto", hl_class="hs-hl")
    with st.container(border=True):
        st.markdown(":gray-badge[:material/visibility_off: Only in the app] Compare the options; the report shows "
                    "only the configuration you choose below.")
        st.html(options)
        st.caption(vsan_options_note(v))
        opts = vsan_pick_options(v)
        ss = st.session_state
        if ss.get("vsan_pick") not in opts:
            ss.vsan_pick = VSAN_PICK_AUTO
        st.selectbox("Configuration for the report", opts, key="vsan_pick", format_func=lambda x: vsan_pick_text(v, x),
                     help="The customer report shows only this configuration and its capacity breakdown. "
                          "Closest fit = least raw capacity that meets the need.")
    if breakdown:
        with st.container(border=True):
            st.html(breakdown)
    st.caption(VSAN_NOTE + f" Compression {v['compression']:.2f}x; up to {v['max_drives']} drives per host.")


def page_recommendation():
    ctx = require_report("Recommendation")
    if not ctx:
        return
    r = ctx["r"]
    page_header("Recommendation", context_subtitle(ctx), ctx)
    scope_bar(ctx)
    if ctx.get("out_of_scope"):
        with st.container(border=True):
            st.markdown("#### This cluster is out of scope")
            st.caption("It isn't sized and won't appear in any report. Turn **Active** on above (or on Current "
                       "Environment) to configure it.")
        return
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
    if d.get("ghz_use") is not None:
        cards.insert(3, ("Measured CPU at failover", f"{d['ghz_use']:.0%}", "of usable GHz",
                         f"{r['perf_ghz']:,.0f} GHz ({r['lo_basis']}) vs {d['usable_ghz']:,.0f} GHz usable",
                         badge_md(d["ghz_ok"], "Fits", "Short on CPU")))
    for col, (label, val, unit, sub, badge) in zip(st.columns(len(cards)), cards):
        with col, st.container(border=True, height="stretch"):
            eyebrow(label)
            big_value(val, unit, sub)
            st.markdown(badge)

    if r["headroom_now"] is not None:
        section_title("Capacity planning", "Room to grow")
        st.caption(f"How many more VMs of today's average size fit, with the HA hosts failed and the "
                   f"{r['vcpu_ratio']:.1f}:1 design limit, before another host is needed.")
        g = st.columns(4)
        g[0].metric("Average VM", f"{r['avg_vm_vcpu']:.1f} vCPU · {r['avg_vm_ram']:,.0f} GB", border=True,
                    help=f"Across {r['tot_vms']:,} VMs in scope today.")
        for col, (label, h, hosts) in zip(g[1:3], (("Day 1 room", r["headroom_now"], r["hosts_now"]),
                                                   ("Future room", r["headroom_fut"], r["hosts_fut"]))):
            big, unit, sub = headroom_view(h, "")
            col.metric(f"{label} · {hosts} hosts", f"{big} {unit}".strip(), border=True, help=sub.strip() or None)
        g[3].metric("Each extra host", f"~{r['vms_per_host']:,} VMs" if r["vms_per_host"] else "n/a", border=True,
                    help="Room one more host adds at the design limit.")
        if r["vcpu_ratio"] > r["hv_ratio"] + EPS:
            st.caption(f":red[The design limit ({r['vcpu_ratio']:.1f}:1) is above the {hv_label(r)}, so this headroom "
                       "assumes a denser cluster than the guidance.]")


    section_title("Side by side", "Current vs. " + ("proposed" if r["override_now"] else "recommended"))
    day1_col = "Proposed (Day 1)" if r["override_now"] else "Recommended (Day 1)"
    fut_col = f"Future ({r['years']} yrs)"
    view = st.segmented_control("Compare", ["Day 1", "Future", "Both"], key="cmp_view",
                                width=seg_width(["Day 1", "Future", "Both"]),
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
            ("CPU", f"{host_cpu_text(r)} · {r['host_cores']} cores", False),
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

    if r["hypervisor"] == "VMware" and r.get("vsan_show"):
        render_vsan(r)
    elif r["sto"].get("vsan_src"):
        st.caption(":material/storage: vSAN was found in the current environment. Turn on **Show vSAN options** under "
                   "Sizing parameters > Storage to estimate a vSAN host configuration.")

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
    bundle = ctx.get("bundle") or {}
    items = bundle.get("items") or []
    if items:
        with st.container(border=True):
            st.markdown("#### What's in the consolidated report")
            st.caption("A Summary tab with totals, a cluster comparison, licensing and storage roll-up and risks, "
                       "then one tab per sized cluster. Consolidation scenarios get their own tab but aren't added to the totals.")
            st.markdown(" ".join(
                f":{'blue' if it['kind'] == 'cluster' else 'violet'}-badge[{it['label']}"
                f"{' · scenario' if it['kind'] == 'scenario' else ''}]" for it in items))
            sized = {c for it in items if it["kind"] == "cluster" for c in [it["label"]]}
            missing = [c for c in ctx["clusters"] if c not in sized]
            if missing and len(ctx["clusters"]) > 1:
                st.caption(f":material/info: Not sized yet, so not included: {', '.join(missing)}. "
                           "Select a cluster on Recommendation to add it.")
    section_title("Preview", "What the customer will see")
    view = "This selection"
    if bundle.get("html"):
        view = st.segmented_control("Preview", ["Consolidated", "This selection"], key="preview_view",
                                    width=seg_width(["Consolidated", "This selection"]),
                                    label_visibility="collapsed") or "Consolidated"
    preview = bundle["html"] if view == "Consolidated" and bundle.get("html") else ctx["html"]
    if hasattr(st, "iframe"):
        st.iframe(preview, height=1400)
    else:
        components.html(preview, height=1400, scrolling=True)


def page_raw():
    ctx = require_report("Raw Data")
    if not ctx:
        return
    page_header("Raw Data", context_subtitle(ctx), ctx)
    c1, c2, c3 = st.columns([1, 1.4, 2], vertical_alignment="bottom")
    view = c1.segmented_control("Dataset", ["VMs", "Hosts"], key="raw_view",
                                width=seg_width(["VMs", "Hosts"])) or DEFAULTS["raw_view"]
    cols_mode = c2.segmented_control("Columns", ["Key columns", "All columns"], key="raw_cols",
                                     width=seg_width(["Key columns", "All columns"])) or DEFAULTS["raw_cols"]
    query = c3.text_input("Search", key="raw_search", icon=":material/search:", placeholder="Filter rows...")

    df = ctx["inv"]["raw_vms"] if view == "VMs" else ctx["inv"]["raw_hosts"]
    if cols_mode == "Key columns":
        wanted = (["Source"] + [c for t in ("RVTools", "LiveOptics") if t in ctx["source"] for c in KEY_COLUMNS[t][view]])
        keep = [c for c in dict.fromkeys(wanted) if c in df.columns]
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
    st.set_page_config(page_title=f"{PRODUCT_NAME} | AHEAD", layout="wide",
                       page_icon=PRODUCT_ICON_PATH if os.path.exists(PRODUCT_ICON_PATH) else ":material/dns:")
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
        ensure_streamlit_credentials()
        print("AHEAD theme config is in place.")
    else:
        main()
