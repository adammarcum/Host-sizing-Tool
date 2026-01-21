import streamlit as st
import pandas as pd
import math
import os
from datetime import datetime

# --- CONFIGURATION ---
APP_TITLE = "Virtualization Sizing Calculator"
DEFAULT_LOGO = "https://placehold.co/200x50/004B87/ffffff?text=AHEAD"

# --- HELPER FUNCTIONS ---
def safe_sum(df, col):
    if col in df.columns:
        return pd.to_numeric(df[col], errors='coerce').sum()
    return 0

def get_rvtools_tb(df, base):
    """
    RVTools Specific: Searches for {base} + Unit (MiB, MB, GB, TB)
    and normalizes to Terabytes.
    """
    if f"{base} MiB" in df.columns: return safe_sum(df, f"{base} MiB") / 1048576
    if f"{base} MB" in df.columns: return safe_sum(df, f"{base} MB") / 1048576
    if f"{base} TB" in df.columns: return safe_sum(df, f"{base} TB")
    if f"{base} GB" in df.columns: return safe_sum(df, f"{base} GB") / 1024
    return 0.0

def calc_license_cores(sockets, cores_per_socket):
    billable_per_socket = max(cores_per_socket, 16)
    return sockets * billable_per_socket

def generate_html_report(data, cluster_name, source_filename, customer_name, logo_url):
    """Generates a full-fidelity HTML report."""
    now = datetime.now().strftime("%Y-%m-%d")
    lic_color = "#d9534f" if data['lic_diff'] > 0 else "#28a745"
    lic_prefix = "+" if data['lic_diff'] > 0 else ""

    # NUMA Logic
    wide_vm_html = ""
    if data['max_vm_cpu'] > 0:
        is_wide_cpu = data['max_vm_cpu'] > data['tgt_numa_cores']
        is_wide_ram = data['max_vm_ram'] > data['tgt_numa_ram']
        if not is_wide_cpu and not is_wide_ram:
            wide_vm_html = "<div style='color:#28a745; font-weight:bold;'>✅ Healthy: All VMs fit within new NUMA boundaries.</div>"
        else:
            if is_wide_cpu: wide_vm_html += f"<div style='color:#d9534f; margin-bottom:5px;'>⚠️ <strong>Wide CPU:</strong> '{data['name_max_cpu']}' ({data['max_vm_cpu']} vCPU) exceeds socket width.</div>"
            if is_wide_ram: wide_vm_html += f"<div style='color:#d9534f;'>⚠️ <strong>Wide RAM:</strong> '{data['name_max_ram']}' ({data['max_vm_ram']:.0f} GB) exceeds socket RAM.</div>"
    else: wide_vm_html = "<div style='color:#666;'>No VM data found.</div>"

    # Performance Section HTML (Live Optics Only)
    perf_html = ""
    if data['has_perf']:
        # Dynamic Insight
        if data['perf_hosts_rec'] < data['hosts_now']:
            insight_title = "💡 Consolidation Opportunity"
            insight_color = "#28a745"
            insight_msg = f"Actual workload ({data['lo_basis']}) is lighter than allocation. You could consolidate to <strong>{data['perf_hosts_rec']} Hosts</strong>."
        else:
            insight_title = "⚠️ Performance Risk"
            insight_color = "#d9534f"
            insight_msg = f"Actual workload ({data['lo_basis']}) demands <strong>{data['perf_hosts_rec']} Hosts</strong>, which is higher than allocation. The environment may be running hot."

        perf_html = f"""
        <h2>6. Performance Analysis (Live Optics)</h2>
        <div class="card">
            <div class="section-label">Allocation vs. Consumption ({data['lo_basis']})</div>
            <div class="grid">
                <div>
                    <div style="font-size:0.9em; color:#666;">Allocated (Entitlement)</div>
                    <div class="metric">{data['tot_vcpu']:,.0f} vCPU</div>
                    <div style="font-size:0.8em;">Total Configured vCPUs</div>
                </div>
                <div>
                    <div style="font-size:0.9em; color:#666;">Consumed ({data['lo_basis']})</div>
                    <div class="metric">{data['perf_ghz_demand']:,.1f} GHz</div>
                    <div style="font-size:0.8em;">Aggregate CPU Demand</div>
                </div>
            </div>
            <div style="margin-top:20px; padding-top:10px; border-top:1px solid #ddd;">
                <div class="section-label" style="color:{insight_color};">{insight_title}</div>
                <div style="font-size:1.1em; margin-bottom:10px;">{insight_msg}</div>
                <div class="note-box">
                    <strong>Basis:</strong> "Safe Sizing" guarantees 100% entitlement performance. "Performance Sizing" relies on the {data['lo_basis']} metric captured during the Live Optics window.
                </div>
            </div>
        </div>
        """

    # Storage Display Logic
    store_infra_html = ""
    if data['ds_cap'] > 0:
        store_infra_html = f"""
        <table>
            <tr><td>Total Capacity:</td><td><strong>{data['ds_cap']:,.1f} TB</strong></td></tr>
            <tr><td>Total In Use:</td><td><strong>{data['ds_used']:,.1f} TB</strong></td></tr>
            <tr><td>Free Space:</td><td><strong>{data['ds_free']:,.1f} TB</strong></td></tr>
        </table>"""
    else:
        store_infra_html = "<div style='color:#999; padding:10px;'><em>Infrastructure storage data not found in source file.</em></div>"

    html = f"""
    <html>
    <head>
        <title>Sizing Report - {customer_name}</title>
        <style>
            body {{ font-family: "Segoe UI", Roboto, Helvetica, Arial, sans-serif; max-width: 1000px; margin: auto; padding: 40px; color: #333; background-color: #fff; }}
            .header-container {{ border-bottom: 3px solid #004B87; padding-bottom: 20px; margin-bottom: 30px; display: flex; justify-content: space-between; align-items: center; }}
            .header-text h1 {{ color: #000; margin: 0; font-size: 24px; text-transform: uppercase; letter-spacing: 1px; }}
            .sub-header {{ color: #666; font-size: 14px; margin-top: 5px; }}
            .header-logo img {{ max-height: 60px; width: auto; }}
            h2 {{ color: #004B87; margin-top: 40px; border-left: 5px solid #004B87; padding-left: 10px; text-transform: uppercase; font-size: 1.1em; page-break-after: avoid; }}
            .card {{ background: #f9f9f9; padding: 20px; border-radius: 4px; border: 1px solid #eee; margin-bottom: 20px; page-break-inside: avoid; }}
            .grid {{ display: grid; grid-template-columns: 1fr 1fr; gap: 20px; }}
            .metric {{ font-size: 1.6em; font-weight: bold; color: #2c3e50; margin: 5px 0; }}
            .sub-metric {{ font-size: 0.9em; color: #666; }}
            .highlight {{ background-color: #e6f3ff; border: 1px solid #b6d4fe; padding: 20px; border-radius: 8px; page-break-inside: avoid; }}
            .section-label {{ font-weight: bold; color: #004B87; margin-bottom: 8px; display: block; text-transform: uppercase; font-size: 0.8em; }}
            table {{ width: 100%; border-collapse: collapse; margin-bottom: 10px; }}
            th, td {{ border-bottom: 1px solid #ddd; padding: 8px; text-align: left; font-size: 0.9em; }}
            th {{ color: #555; text-transform: uppercase; font-size: 0.8em; background: #f1f1f1; }}
            .note-box {{ background-color: #fff3cd; color: #856404; padding: 10px; border-radius: 4px; font-size: 0.85em; margin-bottom: 10px; border: 1px solid #ffeeba; }}
            .lic-delta {{ font-weight: bold; color: {lic_color}; }}
            .footer {{ text-align: center; margin-top: 50px; color: #999; font-size: 0.8em; border-top: 1px solid #eee; padding-top: 20px; }}
        </style>
    </head>
    <body>
        <div class="header-container">
            <div class="header-text">
                <h1>{APP_TITLE}</h1>
                <div class="sub-header">Prepared for <strong>{customer_name}</strong></div>
                <div class="sub-header">Date: {now} | Scope: {cluster_name}</div>
            </div>
            <div class="header-logo">
                <img src="{logo_url}" alt="Logo">
            </div>
        </div>
        
        <h2>1. Executive Sizing Recommendation</h2>
        <div class="grid" style="margin-bottom: 20px;">
            <div class="highlight">
                <div class="section-label">Current Refresh Requirement</div>
                <div class="metric">{data['hosts_now']} Hosts</div>
                <div class="sub-metric">Configuration: N+{data['ha_nodes']}</div>
                <div class="sub-metric">Efficiency: <strong>{data['ratio_now']:.1f}:1</strong> vCPU:pCPU</div>
                
                <div style="margin-top:20px; border-top:1px solid #b6d4fe; padding-top:10px;">
                    <div class="sub-metric" style="color:#000;">Constraint: <strong>{data['constraint']} Bound</strong></div>
                    <div style="font-size:0.8em; color:#666; margin-top:5px;">
                        Logic: Workload requires {data['raw_hosts_cpu']} hosts (CPU) vs {data['raw_hosts_ram']} (RAM). Active: {int(data['raw_hosts'])}. Total: {data['hosts_now']}.
                    </div>
                </div>
            </div>
            <div class="highlight">
                <div class="section-label">Future Requirement with Growth</div>
                <div class="metric">{data['hosts_fut']} Hosts</div>
                <div class="sub-metric">Growth Model: {data['growth']*100:.0f}% Annually</div>
                <div class="sub-metric">Efficiency: <strong>{data['ratio_fut']:.1f}:1</strong> vCPU:pCPU</div>
                <div class="sub-metric">Projected vCPU: {data['fut_vcpu']:,.0f}</div>
            </div>
        </div>

        <div class="card">
            <div class="section-label">Target Hardware Specification</div>
            <div class="grid">
                <div>
                    <div style="font-weight:bold; margin-bottom:5px;">Per Node Config</div>
                    <div>{data['sockets']} Sockets &times; {data['cores']} Cores</div>
                    <div>{data['host_cap_cores']} Physical Cores</div>
                    <div>{data['ram']} GB RAM</div>
                </div>
                <div>
                    <div style="font-weight:bold; margin-bottom:5px;">Cluster Capacity (Day 1)</div>
                    <div>{data['hosts_now'] * data['host_cap_cores']} Total Cores</div>
                    <div>{data['hosts_now'] * data['ram']:,.0f} GB Total RAM</div>
                    <div style="color:#666; font-size:0.9em; margin-top:5px;">Design Limit: {data['ratio_now']:.1f}:1 Ratio</div>
                </div>
            </div>
        </div>

        <h2>2. Workload Scope</h2>
        <div class="card">
            <div class="grid">
                <div>
                    <span class="section-label">Current Host Supply</span>
                    <table>
                        <tr><th>Hosts</th><th>Physical Cores</th><th>Total RAM</th></tr>
                        <tr><td>{data['cur_host_count']}</td><td>{data['cur_cores']:,.0f}</td><td>{data['cur_total_ram_gb']:,.0f} GB</td></tr>
                    </table>
                </div>
                <div>
                    <span class="section-label">VM Demand</span>
                    <table>
                        <tr><th>VMs</th><th>vCPU</th><th>vRAM</th><th>Current Ratio</th></tr>
                        <tr><td>{data['tot_vms']}</td><td>{data['tot_vcpu']:,.0f}</td><td>{data['tot_ram']:,.0f} GB</td><td><strong>{data['cur_ratio']:.1f}:1</strong></td></tr>
                    </table>
                </div>
            </div>
        </div>

        <h2>3. Storage Requirements</h2>
        <div class="grid">
            <div class="card">
                <div class="section-label">VM Allocation (vInfo)</div>
                <div style="font-size:0.85em; color:#666; margin-bottom:10px;">Allocated by VMs</div>
                <table>
                    <tr><td>Provisioned:</td><td><strong>{data['vinfo_prov']:,.1f} TB</strong></td></tr>
                    <tr><td>In Use:</td><td><strong>{data['vinfo_used']:,.1f} TB</strong></td></tr>
                    <tr><td>Backup Scope:</td><td><strong>{data['bak_cons']:,.1f} TB</strong></td></tr>
                </table>
            </div>
            <div class="card">
                <div class="section-label">Infrastructure (Raw)</div>
                <div class="note-box">Note: Capacity represents Global Shared Storage if not filtered.</div>
                {store_infra_html}
            </div>
        </div>

        <h2>4. Architecture & NUMA</h2>
        <div class="grid">
            <div class="card">
                <div class="section-label">NUMA Boundary</div>
                <table>
                    <tr><th>Hardware</th><th>Cores</th><th>Memory</th></tr>
                    <tr><td>Current</td><td>{data['cur_numa_cores_est']:.0f}</td><td>{data['cur_numa_ram_est']:.0f} GB</td></tr>
                    <tr><td>Target</td><td>{data['tgt_numa_cores']}</td><td>{data['tgt_numa_ram']:.0f} GB</td></tr>
                </table>
            </div>
            <div class="card">
                <div class="section-label">Large VM Check</div>
                {wide_vm_html}
            </div>
        </div>

        <h2>5. Licensing Impact</h2>
        <div class="card">
            <div class="grid">
                <div>
                    <div class="section-label">Legacy State</div>
                    <div class="metric">{data['cur_lic_cores']:,.0f} Cores</div>
                    <div style="margin-bottom:10px;"><strong>Edition:</strong> {data['lic_edition']}</div>
                    <div class="sub-metric">{data['cur_host_count']} Hosts</div>
                </div>
                <div>
                    <div class="section-label">Future State</div>
                    <div class="metric">{data['fut_lic_cores']:,.0f} Cores</div>
                    <div class="sub-metric">{data['hosts_fut']} Hosts</div>
                    <div class="lic-delta" style="margin-top:5px;">Net: {lic_prefix}{data['lic_diff']} Cores</div>
                </div>
            </div>
        </div>

        {perf_html}

        <div class="footer">Generated by {APP_TITLE} | Source: {source_filename}</div>
    </body>
    </html>
    """
    return html

# --- PAGE CONFIG ---
st.set_page_config(page_title=APP_TITLE, layout="wide", page_icon="📊")

# --- SIDEBAR ---
st.sidebar.title("⚙️ Sizing Parameters")
st.sidebar.subheader("1. Target Hardware")
tgt_sockets = st.sidebar.number_input("Sockets/Host", 1, 4, 2)
tgt_cores = st.sidebar.number_input("Cores/Socket", 4, 128, 24)
tgt_ram = st.sidebar.number_input("RAM/Host (GB)", 64, 4096, 1024)
tgt_clock = st.sidebar.number_input("CPU Speed (GHz)", 1.0, 5.0, 2.5, help="Used for Live Optics Performance sizing")

lo_basis = st.sidebar.radio(
    "Live Optics Basis", 
    ["95th Percentile", "Peak CPU", "Average CPU"], 
    index=0, 
    help="Determines the GHZ demand used for performance sizing. '95th' is estimated as 95% of Peak if raw data is missing."
)

host_cap_cores = tgt_sockets * tgt_cores
tgt_numa_cores = tgt_cores
tgt_numa_ram = tgt_ram / tgt_sockets
st.sidebar.info(f"**Host Spec:**\n{host_cap_cores} Cores | {tgt_ram} GB RAM\n\n**NUMA Node:**\n{tgt_numa_cores} Cores | {tgt_numa_ram:.0f} GB")

st.sidebar.subheader("2. Constraints")
vcpu_ratio = st.sidebar.slider("Max vCPU:pCPU (Design Limit)", 1.0, 10.0, 5.0, 0.5)
cpu_buffer = st.sidebar.slider("CPU Overhead (%)", 0, 50, 10)
ram_buffer = st.sidebar.slider("RAM Overhead (%)", 0, 50, 10)
min_hosts = st.sidebar.number_input("Min Cluster Size", 1, 32, 3)
ha_nodes = st.sidebar.number_input("HA Tolerance", 0, 4, 1)

st.sidebar.subheader("3. Scope & Growth")
include_off = st.sidebar.checkbox("Include Powered Off VMs?", value=True)
growth = st.sidebar.number_input("Annual Growth (%)", 0.0, 100.0, 10.0) / 100
years = st.sidebar.number_input("Years", 1, 10, 3)

st.sidebar.divider()
st.sidebar.subheader("📥 Report Settings")
cust_name = st.sidebar.text_input("Customer Name", "My Customer")
logo_url = st.sidebar.text_input("Logo URL", DEFAULT_LOGO)
if logo_url: st.sidebar.image(logo_url, width=150)

# --- MAIN ---
st.title(f"📊 {APP_TITLE}")
st.markdown(f"Automated hardware sizing analysis for **{cust_name}**.")

uploaded_files_list = st.file_uploader(
    "Upload Source Files (RVTools .xlsx or Live Optics .xlsx)", 
    type=["xlsx"], 
    accept_multiple_files=True,
    label_visibility="visible"
)

if not uploaded_files_list:
    st.info("👆 Please upload one or more files to begin.")
    st.stop()

# --- FILE SELECTION LOGIC ---
file_map = {f.name: f for f in uploaded_files_list}
selected_filename = st.selectbox("Select File to Analyze:", list(file_map.keys()))
uploaded_file = file_map[selected_filename]

# --- PARSERS ---
def process_rvtools(sheets, selected_cluster, include_off):
    # vInfo
    df_info = sheets['vInfo']
    df_info.columns = df_info.columns.str.strip()
    if 'Cluster' in df_info.columns and selected_cluster != "All Clusters":
        df_info = df_info[df_info['Cluster'] == selected_cluster]
    if 'Powerstate' in df_info.columns and not include_off:
        df_info = df_info[df_info['Powerstate'].astype(str).str.contains('poweredOn', case=False, na=False)]
    
    if 'Memory' in df_info.columns and 'Memory GB' not in df_info.columns:
        df_info['Memory GB'] = pd.to_numeric(df_info['Memory'], errors='coerce') / 1024
    
    # vHost
    cur_cores, cur_host_count, cur_total_ram_gb = 0, 0, 0
    cur_numa_cores_est, cur_numa_ram_est, cur_lic_cores = 0, 0, 0
    lic_edition = "Unknown"
    
    if 'vHost' in sheets:
        df_h = sheets['vHost']
        df_h.columns = df_h.columns.str.strip()
        if 'Cluster' in df_h.columns and selected_cluster != "All Clusters":
            df_h = df_h[df_h['Cluster'] == selected_cluster]
        
        cur_host_count = len(df_h)
        if not df_h.empty:
            df_h['TC'] = df_h['# CPU'] * df_h['Cores per CPU']
            cur_cores = pd.to_numeric(df_h['TC'], errors='coerce').sum()
            cur_total_ram_gb = pd.to_numeric(df_h['# Memory'], errors='coerce').sum() / 1024
            try:
                cur_sockets_mode = df_h['# CPU'].mode()[0]
                cur_cores_mode = df_h['Cores per CPU'].mode()[0]
                cur_ram_mb_mode = df_h['# Memory'].mode()[0]
                cur_numa_cores_est = cur_cores_mode
                cur_numa_ram_est = (cur_ram_mb_mode / 1024) / cur_sockets_mode
            except: pass
            
            # Licensing
            for _, row in df_h.iterrows():
                cur_lic_cores += (row['# CPU'] * max(row['Cores per CPU'], 16))
            
            # Attempt edition read
            if 'Product' in df_h.columns:
                lic_edition = df_h['Product'].mode()[0]

    # vDatastore (RVTools uses smart matching)
    ds = {"cap": 0, "used": 0, "free": 0, "prov": 0, "note": "Global"}
    if 'vDatastore' in sheets:
        df_ds = sheets['vDatastore']
        df_ds.columns = df_ds.columns.str.strip()
        if 'Name' in df_ds.columns: df_ds = df_ds[~df_ds['Name'].astype(str).str.contains('local', case=False, na=False)]
        if 'Cluster' in df_ds.columns and selected_cluster != "All Clusters":
            df_ds = df_ds[df_ds['Cluster'] == selected_cluster]
            ds["note"] = f"Filtered to {selected_cluster}"
        ds["cap"] = get_rvtools_tb(df_ds, 'Capacity')
        ds["used"] = get_rvtools_tb(df_ds, 'In Use')
        ds["prov"] = get_rvtools_tb(df_ds, 'Provisioned')
        ds["free"] = ds["cap"] - ds["used"]

    # Backup & RDM
    bak_cons, rdm_cap, rdm_cnt = 0, 0, 0
    if 'vPartition' in sheets:
        df_p = sheets['vPartition']
        df_p.columns = df_p.columns.str.strip()
        if 'VM' in df_p.columns and 'VM' in df_info.columns: df_p = df_p[df_p['VM'].isin(df_info['VM'])]
        bak_cons = get_rvtools_tb(df_p, 'Consumed')
    
    if 'vDisk' in sheets:
        df_d = sheets['vDisk']
        df_d.columns = df_d.columns.str.strip()
        if 'VM' in df_d.columns and 'VM' in df_info.columns: df_d = df_d[df_d['VM'].isin(df_info['VM'])]
        if 'Raw' in df_d.columns:
            df_rdm = df_d[df_d['Raw'].astype(str).str.contains('True', case=False, na=False)]
            rdm_cnt = len(df_rdm)
            rdm_cap = get_rvtools_tb(df_rdm, 'Capacity')

    # Max VM
    max_vm_cpu, max_vm_ram = 0, 0
    name_max_cpu, name_max_ram = "N/A", "N/A"
    if not df_info.empty:
        max_vm_cpu = df_info['CPUs'].max()
        max_vm_ram = df_info['Memory GB'].max()
        try: name_max_cpu = df_info.loc[df_info['CPUs'].idxmax(), 'VM']
        except: pass
        try: name_max_ram = df_info.loc[df_info['Memory GB'].idxmax(), 'VM']
        except: pass

    return {
        'tot_vms': len(df_info), 'tot_vcpu': safe_sum(df_info, 'CPUs'), 'tot_ram': safe_sum(df_info, 'Memory GB'),
        'vinfo_prov': get_rvtools_tb(df_info, 'Provisioned'), 'vinfo_used': get_rvtools_tb(df_info, 'In Use'),
        'max_vm_cpu': max_vm_cpu, 'max_vm_ram': max_vm_ram, 'name_max_cpu': name_max_cpu, 'name_max_ram': name_max_ram,
        'cur_cores': cur_cores, 'cur_host_count': cur_host_count, 'cur_total_ram_gb': cur_total_ram_gb, 'cur_lic_cores': cur_lic_cores,
        'cur_numa_cores_est': cur_numa_cores_est, 'cur_numa_ram_est': cur_numa_ram_est,
        'ds': ds, 'bak_cons': bak_cons, 'rdm_cap': rdm_cap, 'rdm_cnt': rdm_cnt,
        'has_perf': False, 'perf_ghz_demand': 0, 'lic_edition': lic_edition,
        'df_raw_vinfo': df_info, 'df_raw_vhost': sheets.get('vHost', pd.DataFrame())
    }

def process_live_optics(sheets, selected_cluster, include_off, perf_metric):
    # VMs
    df_vm = sheets['VMs']
    df_vm.columns = df_vm.columns.str.strip()
    if 'Cluster' in df_vm.columns and selected_cluster != "All Clusters":
        df_vm = df_vm[df_vm['Cluster'] == selected_cluster]
    if 'Power State' in df_vm.columns and not include_off:
        df_vm = df_vm[df_vm['Power State'].astype(str).str.contains('poweredOn', case=False, na=False)]

    # Hosts
    df_h = sheets.get('ESX Hosts', pd.DataFrame())
    df_h.columns = df_h.columns.str.strip()
    if 'Cluster' in df_h.columns and selected_cluster != "All Clusters":
        df_h = df_h[df_h['Cluster'] == selected_cluster]

    # Host Devices (Storage) - Use explicit math
    df_dev = sheets.get('Host Devices', pd.DataFrame())
    ds = {"cap": 0, "used": 0, "free": 0, "prov": 0, "note": "Global"}
    if not df_dev.empty:
        df_dev.columns = df_dev.columns.str.strip()
        ds["cap"] = safe_sum(df_dev, 'Capacity (GiB)') / 1024
        ds["used"] = safe_sum(df_dev, 'Used Capacity (GiB)') / 1024
        ds["free"] = safe_sum(df_dev, 'Free Capacity (GiB)') / 1024
        ds["note"] = "Derived from Host Devices (LUNs)"

    # Performance
    has_perf = False
    perf_ghz_demand = 0
    if 'ESX Performance' in sheets:
        df_perf = sheets['ESX Performance']
        df_perf.columns = df_perf.columns.str.strip()
        if 'Cluster' in df_perf.columns and selected_cluster != "All Clusters":
            df_perf = df_perf[df_perf['Cluster'] == selected_cluster]
        
        # 95th Percentile Logic
        if perf_metric == "95th Percentile":
            if '95th Percentile CPU (GHz)' in df_perf.columns:
                perf_ghz_demand = safe_sum(df_perf, '95th Percentile CPU (GHz)')
            else:
                perf_ghz_demand = safe_sum(df_perf, 'Peak CPU (GHz)') * 0.95
        elif perf_metric == "Peak CPU":
            perf_ghz_demand = safe_sum(df_perf, 'Peak CPU (GHz)')
        else: # Average
            perf_ghz_demand = safe_sum(df_perf, 'Average CPU (GHz)')
            
        has_perf = True

    # Host Supply Logic
    cur_cores = safe_sum(df_h, 'CPU Cores')
    cur_total_ram_gb = safe_sum(df_h, 'Memory (KiB)') / 1024 / 1024
    cur_lic_cores = 0
    cur_numa_cores_est, cur_numa_ram_est = 0, 0
    lic_edition = "Unknown"
    
    if not df_h.empty:
        try:
            cur_sockets_mode = df_h['CPU Sockets'].mode()[0]
            df_h['CoresPerSocket'] = df_h['CPU Cores'] / df_h['CPU Sockets']
            cur_cores_mode = df_h['CoresPerSocket'].mode()[0]
            cur_numa_cores_est = cur_cores_mode
            cur_numa_ram_est = (cur_total_ram_gb / len(df_h)) / cur_sockets_mode
        except: pass
        
        for _, row in df_h.iterrows():
            try: cur_lic_cores += (row['CPU Sockets'] * max(row['CPU Cores'] / row['CPU Sockets'], 16))
            except: pass

    # License Edition (LO specific)
    if 'ESX Licenses' in sheets:
        df_lic = sheets['ESX Licenses']
        df_lic.columns = df_lic.columns.str.strip()
        if 'Software Title' in df_lic.columns:
            try: lic_edition = df_lic['Software Title'].mode()[0]
            except: pass

    # Max VM
    max_vm_cpu, max_vm_ram = 0, 0
    name_max_cpu, name_max_ram = "N/A", "N/A"
    bak_cons = 0
    
    if not df_vm.empty:
        max_vm_cpu = df_vm['Virtual CPU'].max()
        max_vm_ram = df_vm['Provisioned Memory (MiB)'].max() / 1024
        bak_cons = safe_sum(df_vm, 'Guest VM Disk Used (MiB)') / 1048576 # MiB to TB
        
        try: name_max_cpu = df_vm.loc[df_vm['Virtual CPU'].idxmax(), 'VM Name']
        except: pass
        try: name_max_ram = df_vm.loc[df_vm['Provisioned Memory (MiB)'].idxmax(), 'VM Name']
        except: pass

    return {
        'tot_vms': len(df_vm), 'tot_vcpu': safe_sum(df_vm, 'Virtual CPU'), 'tot_ram': safe_sum(df_vm, 'Provisioned Memory (MiB)') / 1024,
        'vinfo_prov': safe_sum(df_vm, 'Virtual Disk Size (MiB)') / 1048576, 'vinfo_used': safe_sum(df_vm, 'Virtual Disk Used (MiB)') / 1048576,
        'max_vm_cpu': max_vm_cpu, 'max_vm_ram': max_vm_ram, 'name_max_cpu': name_max_cpu, 'name_max_ram': name_max_ram,
        'cur_cores': cur_cores, 'cur_host_count': len(df_h), 'cur_total_ram_gb': cur_total_ram_gb, 'cur_lic_cores': cur_lic_cores,
        'cur_numa_cores_est': cur_numa_cores_est, 'cur_numa_ram_est': cur_numa_ram_est,
        'ds': ds, 'bak_cons': bak_cons, 'rdm_cap': 0, 'rdm_cnt': 0,
        'has_perf': has_perf, 'perf_ghz_demand': perf_ghz_demand, 'lic_edition': lic_edition,
        'df_raw_vinfo': df_vm, 'df_raw_vhost': df_h
    }

# --- EXECUTION ---
try:
    sheets = pd.read_excel(uploaded_file, sheet_name=None, engine='openpyxl')
    
    if 'vInfo' in sheets:
        source_type = "RVTools"
        cluster_source = sheets['vInfo']
        cluster_col = 'Cluster'
    elif 'VMs' in sheets and 'ESX Hosts' in sheets:
        source_type = "LiveOptics"
        cluster_source = sheets['VMs']
        cluster_col = 'Cluster'
    else:
        st.error("Unknown file. Upload RVTools or Live Optics Excel.")
        st.stop()

    st.success(f"📂 **{source_type}** Detected")
    
    if cluster_col and cluster_col in cluster_source.columns:
        clusters = sorted(cluster_source[cluster_col].dropna().unique())
        selected_cluster = st.selectbox("Select Cluster Scope:", clusters)
    else: selected_cluster = "All Clusters"

    if source_type == "RVTools":
        db = process_rvtools(sheets, selected_cluster, include_off)
    else:
        db = process_live_optics(sheets, selected_cluster, include_off, lo_basis)

    eff_host_cores = host_cap_cores * (1 - cpu_buffer/100)
    eff_host_ram = tgt_ram * (1 - ram_buffer/100)
    
    hosts_for_cpu = math.ceil(db['tot_vcpu'] / vcpu_ratio / eff_host_cores)
    hosts_for_ram = math.ceil(db['tot_ram'] / eff_host_ram)
    
    if hosts_for_cpu > hosts_for_ram:
        constraint = "CPU"
        raw_hosts = hosts_for_cpu
        raw_hosts_cpu = hosts_for_cpu
        raw_hosts_ram = hosts_for_ram
    else:
        constraint = "RAM"
        raw_hosts = hosts_for_ram
        raw_hosts_cpu = hosts_for_cpu
        raw_hosts_ram = hosts_for_ram

    mult = (1 + growth) ** years
    fut_vcpu = db['tot_vcpu'] * mult
    fut_hosts_for_cpu = math.ceil(fut_vcpu / vcpu_ratio / eff_host_cores)
    fut_hosts_for_ram = math.ceil((db['tot_ram'] * mult) / eff_host_ram)
    fut_raw_hosts = max(fut_hosts_for_cpu, fut_hosts_for_ram)
    
    hosts_now = max(int(raw_hosts) + ha_nodes, min_hosts)
    hosts_fut = max(int(fut_raw_hosts) + ha_nodes, min_hosts)
    
    pcores_now = hosts_now * host_cap_cores
    cur_ratio = db['tot_vcpu'] / db['cur_cores'] if db['cur_cores'] > 0 else 0
    ratio_now = db['tot_vcpu'] / pcores_now
    ratio_fail = db['tot_vcpu'] / ((hosts_now - ha_nodes) * host_cap_cores) if (hosts_now - ha_nodes) > 0 else 0
    ratio_fut = fut_vcpu / (hosts_fut * host_cap_cores)

    perf_hosts_rec = 0
    if db['has_perf']:
        node_ghz_cap = tgt_sockets * tgt_cores * tgt_clock
        node_ghz_usable = node_ghz_cap * (1 - cpu_buffer/100)
        hosts_needed_ghz = math.ceil(db['perf_ghz_demand'] / (node_ghz_usable * 0.8))
        perf_hosts_rec = max(hosts_needed_ghz + ha_nodes, min_hosts)

    fut_lic_per_node = calc_license_cores(tgt_sockets, tgt_cores)
    fut_lic_cores = hosts_fut * fut_lic_per_node
    lic_diff = fut_lic_cores - db['cur_lic_cores']
    lic_prefix = "+" if lic_diff > 0 else ""

    report_data = {
        'hosts_now': hosts_now, 'ha_nodes': ha_nodes, 'ratio_now': ratio_now,
        'hosts_fut': hosts_fut, 'years': years, 'growth': growth, 'fut_vcpu': fut_vcpu,
        'tot_vms': db['tot_vms'], 'tot_vcpu': db['tot_vcpu'], 'tot_ram': db['tot_ram'],
        'sockets': tgt_sockets, 'cores': tgt_cores, 'ram': tgt_ram, 'host_cap_cores': host_cap_cores,
        'ds_cap': db['ds']['cap'], 'ds_used': db['ds']['used'], 'ds_free': db['ds']['free'], 'ds_scope_note': db['ds']['note'],
        'vinfo_prov': db['vinfo_prov'], 'vinfo_used': db['vinfo_used'], 'bak_cons': db['bak_cons'],
        'cur_host_count': db['cur_host_count'], 'cur_lic_cores': int(db['cur_lic_cores']),
        'fut_lic_cores': int(fut_lic_cores), 'lic_diff': int(lic_diff),
        'cur_cores': db['cur_cores'], 'cur_total_ram_gb': db['cur_total_ram_gb'],
        'constraint': constraint, 'raw_hosts_cpu': raw_hosts_cpu, 'raw_hosts_ram': raw_hosts_ram, 'raw_hosts': raw_hosts,
        'cur_numa_cores_est': db['cur_numa_cores_est'], 'cur_numa_ram_est': db['cur_numa_ram_est'],
        'tgt_numa_cores': tgt_numa_cores, 'tgt_numa_ram': tgt_numa_ram,
        'max_vm_cpu': db['max_vm_cpu'], 'max_vm_ram': db['max_vm_ram'],
        'name_max_cpu': db['name_max_cpu'], 'name_max_ram': db['name_max_ram'],
        'cur_ratio': cur_ratio, 'ratio_fut': ratio_fut,
        'has_perf': db['has_perf'], 'perf_ghz_demand': db['perf_ghz_demand'], 'perf_hosts_rec': perf_hosts_rec, 'lo_basis': lo_basis,
        'lic_edition': db['lic_edition']
    }

    clean_filename = os.path.splitext(uploaded_file.name)[0]
    out_name = f"{clean_filename}_{selected_cluster}_Sizing.html"
    html_string = generate_html_report(report_data, selected_cluster, uploaded_file.name, cust_name, logo_url)
    st.sidebar.download_button("Download Report", html_string, file_name=out_name, mime="text/html")

    tab1, tab2 = st.tabs(["📋 Executive Report", "🔍 Raw Data Analysis"])

    with tab1:
        st.subheader("1. Executive Sizing Recommendation")
        c1, c2 = st.columns(2)
        with c1:
            st.info(f"### 📅 Current Refresh Requirement", icon="📅")
            st.markdown(f"**Cluster Config:** {hosts_now} Nodes (N+{ha_nodes} HA)")
            
            st.markdown("#### 🔧 Hardware Specs")
            st.write(f"**Per Node:** {tgt_sockets} Sockets | {host_cap_cores} Cores | {tgt_ram} GB RAM")
            st.write(f"**Cluster Total:** {hosts_now * host_cap_cores} Cores | {hosts_now * tgt_ram:,.0f} GB RAM")
            st.markdown("#### 📉 CPU Oversubscription")
            st.write(f"**Active Ratio:** {ratio_fail:.1f}:1")
            
            st.divider()
            if constraint == "RAM": st.warning(f"⚠️ **Constraint: Memory Bound**")
            else: st.success(f"✅ **Constraint: CPU Bound**")
            
            with st.expander("📊 Sizing Logic"):
                st.write(f"**1. Workload:** {db['tot_vcpu']:,.0f} vCPU, {db['tot_ram']:,.0f} GB RAM")
                st.write(f"**2. Effective Host:** {eff_host_cores:.1f} Cores, {eff_host_ram:.1f} GB RAM")
                st.write(f"**3. Hosts Needed:** CPU: {hosts_for_cpu}, RAM: {hosts_for_ram}")
                st.write(f"**4. Constraint:** {constraint} -> {raw_hosts} active nodes")
                st.write(f"**5. Final:** {raw_hosts} + {ha_nodes} HA = {hosts_now} Hosts")

        with c2:
            st.success(f"### 🚀 Future Requirement with Growth", icon="🚀")
            st.markdown(f"**Cluster Config:** {hosts_fut} Nodes (+{growth*100:.0f}% Growth / {years} Years)")
            st.markdown("#### 🔧 Hardware Specs")
            st.write(f"**Per Node:** {tgt_sockets} Sockets | {host_cap_cores} Cores | {tgt_ram} GB RAM")
            st.write(f"**Cluster Total:** {hosts_fut * host_cap_cores} Cores | {hosts_fut * tgt_ram:,.0f} GB RAM")
            st.markdown("#### 📉 CPU Density")
            st.write(f"**Future Ratio:** {ratio_fut:.1f}:1")

        st.subheader(f"2. Workload Scope ({selected_cluster})")
        with st.container(border=True):
            sc1, sc2, sc3, sc4 = st.columns(4)
            sc1.metric("VMs", f"{db['tot_vms']}")
            sc2.metric("vCPU", f"{db['tot_vcpu']:,.0f}")
            sc3.metric("vRAM", f"{db['tot_ram']:,.0f} GB")
            sc4.metric("Current Ratio", f"{cur_ratio:.1f}:1")

        st.subheader("3. Storage Requirements")
        c_alloc, c_infra = st.columns(2)
        with c_alloc:
            with st.container(border=True):
                st.markdown(f"#### 📦 VM Allocation (VMDK)")
                st.write(f"**Provisioned:** {db['vinfo_prov']:,.1f} TB")
                st.write(f"**In Use:** {db['vinfo_used']:,.1f} TB")
                st.markdown("---")
                st.write(f"🔹 **Guest OS (Backup):** {db['bak_cons']:.1f} TB")
        with c_infra:
            with st.container(border=True):
                st.markdown(f"#### 🏢 Infrastructure ({db['ds']['note']})")
                if db['ds']['cap'] > 0:
                    st.write(f"**Total Capacity:** {db['ds']['cap']:,.1f} TB")
                    st.write(f"**Free Space:** {db['ds']['free']:,.1f} TB")
                else:
                    st.write("Data not available in source file.")

        st.subheader("4. Architecture & NUMA")
        st.write(f"**Target NUMA:** {tgt_numa_cores} Cores | {tgt_numa_ram:.0f} GB RAM")
        if db['max_vm_cpu'] > 0 and (db['max_vm_cpu'] > tgt_numa_cores or db['max_vm_ram'] > tgt_numa_ram):
            st.warning(f"⚠️ Wide VM Detected: {db['name_max_cpu']} ({db['max_vm_cpu']} vCPU, {db['max_vm_ram']:.0f} GB RAM)")
        else:
            st.success("✅ All VMs fit within NUMA")

        st.subheader("5. Licensing")
        st.write(f"**Current Edition:** {db['lic_edition']}")
        st.write(f"**Net Change:** {lic_prefix}{lic_diff:.0f} Cores")

        if db['has_perf']:
            st.divider()
            st.subheader("6. Performance Analysis (Live Optics)")
            p1, p2 = st.columns(2)
            p1.metric("Allocated vCPU", f"{db['tot_vcpu']:,.0f}")
            p2.metric(f"Consumed {lo_basis.split()[0]} GHz", f"{db['perf_ghz_demand']:,.1f} GHz")
            
            perf_insight_color = "green" if perf_hosts_rec < hosts_now else "red"
            st.caption(f"Based on {lo_basis} metrics.")
            st.markdown(f":{perf_insight_color}[**Analysis:** Workload could run on {perf_hosts_rec} Hosts (vs {hosts_now} Allocation).]")

    with tab2:
        st.write("Source Data Preview")
        st.dataframe(db['df_raw_vinfo'].head())

except Exception as e:
    st.error(f"Error processing file: {e}")