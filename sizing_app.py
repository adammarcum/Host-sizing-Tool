import streamlit as st
import pandas as pd
import math
import os
import traceback
from datetime import datetime

# --- 1. PAGE CONFIG (MUST BE FIRST) ---
APP_TITLE = "Virtualization Sizing Calculator"
DEFAULT_LOGO = "https://placehold.co/200x50/004B87/ffffff?text=AHEAD"

st.set_page_config(
    page_title=APP_TITLE, 
    layout="wide", 
    page_icon="📊"
)

# --- UTILS ---
def to_float(series):
    """Safely converts a series to float, handling commas and strings."""
    if series is None: return 0.0
    return pd.to_numeric(series.astype(str).str.replace(',', ''), errors='coerce').fillna(0)

def clean_sheet_names(sheets):
    return {k.strip(): v for k, v in sheets.items()}

def promote_header(df, keywords):
    """Scans for a header row containing ANY of the specific keywords."""
    df = df.reset_index(drop=True)
    # Scan first 20 rows
    for i in range(min(20, len(df))):
        row_str = " ".join(df.iloc[i].astype(str).fillna('').values).lower()
        if any(k.lower() in row_str for k in keywords):
            df.columns = df.iloc[i]
            # Force all columns to strings to prevent NaN header crashes
            df.columns = df.columns.fillna('').astype(str).str.strip()
            df = df[i+1:].reset_index(drop=True)
            return df
    # Fallback
    df.columns = df.iloc[0]
    df.columns = df.columns.fillna('').astype(str).str.strip()
    df = df[1:].reset_index(drop=True)
    return df

def get_col(df, keywords):
    """Finds a column name: Prioritizes EXACT match, then falls back to PARTIAL match."""
    if df is None or df.empty: return None
    if isinstance(keywords, str): keywords = [keywords]
    
    # 1. Exact Match (Case Insensitive & Stripped)
    for kw in keywords:
        exact = next((c for c in df.columns if str(c).strip().lower() == kw.lower()), None)
        if exact: return exact
    
    # 2. Keyword-specific Safety Rules (Prevents "Cluster rule(s)" bug)
    for kw in keywords:
        if kw.lower() == 'cluster':
            known_vars = ['cluster', 'cluster name', 'host cluster', 'vdatastoreclustername', 'vinfocluster']
            for c in df.columns:
                if str(c).strip().lower() in known_vars:
                    return c
            for c in df.columns:
                cl = str(c).lower()
                if 'cluster' in cl and not any(bad in cl for bad in ['rule', 'capacity', 'free', 'space', 'id']):
                    return c
    
    # 3. Partial Match (Case Insensitive)
    for kw in keywords:
        for c in df.columns:
            cl = str(c).lower()
            if kw.lower() in cl:
                if kw.lower() == 'capacity' and 'cluster' in cl: continue
                return c
                
    return None

def safe_sum(df, col_name):
    if not col_name or col_name not in df.columns: return 0.0
    return to_float(df[col_name]).sum()

def get_rvtools_tb(df, col_name):
    """Smart unit finder for RVTools based on column name."""
    if not col_name or col_name not in df.columns: return 0.0
    
    cl = str(col_name).lower()
    val = to_float(df[col_name]).sum()
    
    if 'tb' in cl or 'tib' in cl: return val
    if 'gb' in cl or 'gib' in cl: return val / 1024
    
    # Default for RVTools exports (Capacity MiB, vDatastoreCapacity) is always MiB
    return val / 1048576

def calc_license_cores(sockets, cores_per_socket):
    billable_per_socket = max(cores_per_socket, 16)
    return sockets * billable_per_socket

# --- REPORT GENERATOR ---
def generate_html_report(data, scope_name, source_filename, customer_name, logo_url):
    now = datetime.now().strftime("%Y-%m-%d")
    lic_prefix = "+" if data.get('lic_diff', 0) > 0 else ""
    lic_color = "#d9534f" if data.get('lic_diff', 0) > 0 else "#28a745"

    # NUMA Display
    vm_check_html = "<div style='color:#666;'>No VM data found.</div>"
    if data.get('max_vm_cpu', 0) > 0:
        cpu_stat = "⚠️ Exceeds Socket" if data['max_vm_cpu'] > data['tgt_numa_cores'] else "✅ Fits NUMA"
        ram_stat = "⚠️ Exceeds Socket" if data['max_vm_ram'] > data['tgt_numa_ram'] else "✅ Fits NUMA"
        vm_check_html = f"""
        <div style="font-size:0.9em; margin-bottom:5px;"><strong>{data['name_max_cpu']}</strong>: {data['max_vm_cpu']} vCPU ({cpu_stat})</div>
        <div style="font-size:0.9em;"><strong>{data['name_max_ram']}</strong>: {data['max_vm_ram']:.0f} GB RAM ({ram_stat})</div>"""

    # vSAN Note
    vsan_html = ""
    if data.get('vsan_detected'):
        if data.get('vsan_raw_tib', 0) > 0:
            vsan_html = f"""
            <div style="margin-top:10px; padding:8px; background:#e8f5e9; border:1px solid #c8e6c9; color:#2e7d32; border-radius:4px; font-size:0.9em;">
                <strong>✅ vSAN/VxRail Detected in Source</strong><br>
                Raw Capacity: <strong>{data['vsan_raw_tib']:,.1f} TiB</strong><br>
                <i>Note: Raw TiB is used for VVF/VCF licensing. Usable capacity depends on RAID policy.</i>
            </div>"""
        else:
            vsan_html = """
            <div style="margin-top:10px; padding:8px; background:#e8f5e9; border:1px solid #c8e6c9; color:#2e7d32; border-radius:4px; font-size:0.9em;">
                <strong>✅ vSAN/VxRail Detected in Source</strong><br>
                <i>Note: Storage capacity planning for vSAN depends on RAID policy (RAID1/5/6) and is not calculated here.</i>
            </div>"""

    # Performance Display
    perf_html = ""
    if data.get('has_perf'):
        perf_html = f"""
        <h2>6. Performance Analysis (Live Optics)</h2>
        <div class="card">
            <div class="section-label">Allocation vs. Consumption ({data['lo_basis']})</div>
            <div class="grid">
                <div>
                    <div style="font-size:0.9em; color:#666;">Allocated (Entitlement)</div>
                    <div class="metric">{data['tot_vcpu']:,.0f} vCPU</div>
                </div>
                <div>
                    <div style="font-size:0.9em; color:#666;">Consumed ({data['lo_basis']})</div>
                    <div class="metric">{data['perf_ghz_demand']:,.1f} GHz</div>
                </div>
            </div>
            <div style="margin-top:10px; padding-top:10px; border-top:1px solid #ddd;">
                <div style="font-size:1.1em; margin-bottom:10px;">Workload requires <strong>{data['perf_hosts_rec']} Hosts</strong> to satisfy demand.</div>
            </div>
        </div>"""

    html = f"""
    <html>
    <head>
        <title>Sizing Report - {customer_name}</title>
        <style>
            body {{ font-family: "Segoe UI", sans-serif; max-width: 1000px; margin: auto; padding: 40px; color: #333; background: #fff; }}
            .header-container {{ border-bottom: 3px solid #004B87; padding-bottom: 20px; margin-bottom: 30px; display: flex; justify-content: space-between; align-items: center; }}
            .header-text h1 {{ margin: 0; font-size: 24px; color: #000; text-transform: uppercase; letter-spacing: 1px; }}
            .header-logo img {{ max-height: 60px; }}
            h2 {{ color: #004B87; border-left: 5px solid #004B87; padding-left: 10px; margin-top: 40px; text-transform: uppercase; font-size: 1.1em; }}
            .card {{ background: #f9f9f9; padding: 20px; border-radius: 4px; border: 1px solid #eee; margin-bottom: 20px; page-break-inside: avoid; }}
            .grid {{ display: grid; grid-template-columns: 1fr 1fr; gap: 20px; }}
            .metric {{ font-size: 1.6em; font-weight: bold; color: #2c3e50; margin: 5px 0; }}
            .section-label {{ font-weight: bold; color: #004B87; text-transform: uppercase; font-size: 0.8em; display:block; margin-bottom: 8px; }}
            
            /* FIXED TABLE STYLING */
            table {{ width: 100%; border-collapse: collapse; margin-bottom: 10px; table-layout: fixed; }}
            th, td {{ border-bottom: 1px solid #ddd; padding: 8px; text-align: left; font-size: 0.9em; overflow: hidden; white-space: nowrap; text-overflow: ellipsis; }}
            th {{ background-color: #f1f1f1; color: #555; text-transform: uppercase; font-size: 0.8em; width: 40%; }}
            td {{ width: 60%; }}
            
            .lic-delta {{ font-weight: bold; color: {lic_color}; }}
            .footer {{ margin-top:50px; text-align:center; color:#999; font-size:0.8em; border-top: 1px solid #eee; padding-top: 20px; }}
        </style>
    </head>
    <body>
        <div class="header-container">
            <div class="header-text">
                <h1>{APP_TITLE}</h1>
                <div style="color:#666; font-size:14px;">Prepared for <strong>{customer_name}</strong> | {now}</div>
                <div style="color:#666; font-size:12px; margin-top:5px;">Scope: {scope_name}</div>
            </div>
            <div class="header-logo"><img src="{logo_url}"></div>
        </div>
        
        <h2>1. Executive Sizing Recommendation</h2>
        <div class="grid">
            <div class="card" style="background-color: #e6f3ff; border: 1px solid #b6d4fe;">
                <div class="section-label">Current Refresh Requirement</div>
                <div class="metric">{data['hosts_now']} Hosts</div>
                <div>Configuration: N+{data['ha_nodes']}</div>
                <div>Efficiency: <strong>{data['ratio_now']:.1f}:1</strong> vCPU:pCPU</div>
                <div style="margin-top:10px; font-size:0.9em;">Constraint: <strong>{data['constraint']} Bound</strong></div>
            </div>
            <div class="card" style="background-color: #e6f3ff; border: 1px solid #b6d4fe;">
                <div class="section-label">Future Requirement</div>
                <div class="metric">{data['hosts_fut']} Hosts</div>
                <div>Growth Model: {data['growth']*100:.0f}% Annually over {data['years']} Years</div>
                <div>Projected vCPU: {data['fut_vcpu']:,.0f}</div>
            </div>
        </div>

        <div class="card">
            <div class="section-label">Target Hardware Specification</div>
            <div class="grid">
                <div>
                    <div><strong>Per Node Config</strong></div>
                    <div>{data['sockets']} Sockets x {data['cores']} Cores</div>
                    <div>{data['host_cap_cores']} Physical Cores</div>
                    <div>{data['ram']} GB RAM</div>
                </div>
                <div>
                    <div><strong>Cluster Capacity (Day 1)</strong></div>
                    <div>{data['hosts_now'] * data['host_cap_cores']} Total Cores</div>
                    <div>{data['hosts_now'] * data['ram']:,.0f} GB Total RAM</div>
                </div>
            </div>
        </div>

        <h2>2. Workload Scope</h2>
        <div class="card">
            <div class="grid">
                <div>
                    <span class="section-label">Legacy Supply (Consolidated)</span>
                    <table>
                        <tr><th>Hosts</th><td>{data['cur_host_count']}</td></tr>
                        <tr><th>Physical Cores</th><td>{data['cur_cores']:,.0f}</td></tr>
                        <tr><th>Total RAM</th><td>{data['cur_total_ram_gb']:,.0f} GB</td></tr>
                    </table>
                </div>
                <div>
                    <span class="section-label">VM Demand</span>
                    <table>
                        <tr><th>VMs</th><td>{data['tot_vms']}</td></tr>
                        <tr><th>vCPU</th><td>{data['tot_vcpu']:,.0f}</td></tr>
                        <tr><th>vRAM</th><td>{data['tot_ram']:,.0f} GB</td></tr>
                        <tr><th>Current Ratio</th><td><strong>{data['cur_ratio']:.1f}:1</strong></td></tr>
                    </table>
                </div>
            </div>
        </div>

        <h2>3. Storage Requirements</h2>
        <div class="grid">
            <div class="card">
                <div class="section-label">VM Allocation (vInfo)</div>
                <table>
                    <tr><th>Provisioned</th><td><strong>{data.get('vinfo_prov', 0):,.1f} TB</strong></td></tr>
                    <tr><th>In Use</th><td><strong>{data.get('vinfo_used', 0):,.1f} TB</strong></td></tr>
                </table>
            </div>
            <div class="card">
                <div class="section-label">Infrastructure</div>
                <table>
                    <tr><th>Total Capacity</th><td><strong>{data['ds_cap']:,.1f} TB</strong></td></tr>
                    <tr><th>Free Space</th><td><strong>{data['ds_free']:,.1f} TB</strong></td></tr>
                </table>
                {vsan_html}
            </div>
        </div>

        <h2>4. Architecture & NUMA</h2>
        <div class="grid">
            <div class="card">
                <div class="section-label">NUMA Boundary</div>
                <table>
                    <tr><th>Hardware</th><th>Cores</th><th>Memory</th></tr>
                    <tr><td>Target</td><td>{data['tgt_numa_cores']}</td><td>{data['tgt_numa_ram']:.0f} GB</td></tr>
                </table>
            </div>
            <div class="card">
                <div class="section-label">Large VM Check</div>
                {vm_check_html}
            </div>
        </div>

        <h2>5. Licensing Impact</h2>
        <div class="card">
            <div class="grid" style="grid-template-columns: 1fr 1fr 1fr;">
                <div>
                    <div class="section-label">Legacy State</div>
                    <div class="metric">{data['cur_lic_cores']:,.0f} Cores</div>
                    <div style="margin-bottom:10px;">{data['cur_host_count']} Hosts</div>
                </div>
                <div>
                    <div class="section-label">Current Refresh</div>
                    <div class="metric">{data['now_lic_cores']:,.0f} Cores</div>
                    <div style="margin-bottom:10px;">{data['hosts_now']} Hosts</div>
                </div>
                <div>
                    <div class="section-label">Future w/ Growth</div>
                    <div class="metric">{data['fut_lic_cores']:,.0f} Cores</div>
                    <div class="lic-delta">Net: {lic_prefix}{data['lic_diff']} Cores</div>
                </div>
            </div>
        </div>

        {perf_html}
        
        <div class="footer">Generated by {APP_TITLE} | Source: {source_filename}</div>
    </body>
    </html>
    """
    return html

# --- PROCESSORS ---
def process_data(sheets, source_type, selected_clusters, include_off, perf_metric):
    db = {
        'tot_vms': 0, 'tot_vcpu': 0, 'tot_ram': 0, 
        'vinfo_prov': 0, 'vinfo_used': 0, 'bak_cons': 0,
        'cur_cores': 0, 'cur_host_count': 0, 'cur_total_ram_gb': 0, 'cur_lic_cores': 0,
        'max_vm_cpu': 0, 'max_vm_ram': 0, 'name_max_cpu': "N/A", 'name_max_ram': "N/A",
        'ds_cap': 0, 'ds_used': 0, 'ds_free': 0,
        'has_perf': False, 'perf_ghz_demand': 0, 'lic_edition': "Unknown",
        'vsan_detected': False, 'vsan_raw_tib': 0
    }

    def apply_filter(df, col_c, col_p):
        if df is None: return None
        
        if col_c and selected_clusters:
            if "All Clusters" not in selected_clusters:
                c_name = get_col(df, [col_c, f'vInfo{col_c}', f'vDatastore{col_c}', f'vHost{col_c}'])
                if c_name: 
                    df[c_name] = df[c_name].fillna("Unclustered")
                    df = df[df[c_name].isin(selected_clusters)]
                    
        if col_p and not include_off:
            p_name = get_col(df, [col_p, f'vInfo{col_p}'])
            if p_name: df = df[df[p_name].astype(str).str.contains('poweredOn', case=False, na=False)]
        
        return df

    # 1. VM Processing
    df_vm = None
    if source_type == "RVTools":
        df_vm = promote_header(sheets['vInfo'], ["VM", "vInfoName"])
        df_vm = apply_filter(df_vm, "Cluster", "Powerstate")
        
        cpu_col = get_col(df_vm, ['CPUs', 'vInfoCPUs'])
        ram_col = get_col(df_vm, ['Memory', 'vInfoMemory'])
        prov_col = get_col(df_vm, ['Provisioned MiB', 'vInfoProvisioned', 'Provisioned'])
        used_col = get_col(df_vm, ['In Use MiB', 'vInfoInUse', 'In Use'])
        
        if cpu_col: db['tot_vcpu'] = to_float(df_vm[cpu_col]).sum()
        if ram_col: db['tot_ram'] = get_rvtools_tb(df_vm, ram_col) * 1024 # Converts back to GB
        if prov_col: db['vinfo_prov'] = get_rvtools_tb(df_vm, prov_col)
        if used_col: db['vinfo_used'] = get_rvtools_tb(df_vm, used_col)
        
    elif source_type == "LiveOptics":
        df_vm = promote_header(sheets['VMs'], ["VM Name", "Guest Hostname"])
        df_vm = apply_filter(df_vm, "Cluster", "Power State")
        
        cpu_col = get_col(df_vm, ['Virtual CPU'])
        ram_col = get_col(df_vm, ['Provisioned Memory (MiB)', 'Provisioned Memory'])
        prov_col = get_col(df_vm, ['Virtual Disk Size (MiB)', 'Virtual Disk Size'])
        used_col = get_col(df_vm, ['Virtual Disk Used (MiB)', 'Virtual Disk Used'])
        
        if cpu_col: db['tot_vcpu'] = to_float(df_vm[cpu_col]).sum()
        if ram_col: db['tot_ram'] = to_float(df_vm[ram_col]).sum() / 1024
        if prov_col: db['vinfo_prov'] = to_float(df_vm[prov_col]).sum() / 1048576
        if used_col: db['vinfo_used'] = to_float(df_vm[used_col]).sum() / 1048576

    if df_vm is not None:
        db['tot_vms'] = len(df_vm)
        c_col = cpu_col
        m_col = ram_col
        n_col = get_col(df_vm, ['VM', 'Name', 'vInfoName', 'VM Name'])
        
        if c_col: 
            db['max_vm_cpu'] = to_float(df_vm[c_col]).max()
            if n_col: db['name_max_cpu'] = df_vm.loc[to_float(df_vm[c_col]).idxmax(), n_col]
        if m_col: 
            raw_max_ram = to_float(df_vm[m_col]).max()
            # If RVTools/LiveOptics, max ram is usually in MB/MiB
            db['max_vm_ram'] = raw_max_ram / 1024 if raw_max_ram > 1000 else raw_max_ram
            if n_col: db['name_max_ram'] = df_vm.loc[to_float(df_vm[m_col]).idxmax(), n_col]

    # 2. Host Processing
    df_h = None
    host_to_cluster = {}
    
    if source_type == "RVTools" and 'vHost' in sheets:
        df_h = promote_header(sheets['vHost'], ["Host", "vHostName"])
        
        # Build map before filtering
        nm_col = get_col(df_h, ['Host', 'vHostName'])
        cl_col = get_col(df_h, ['Cluster', 'vHostCluster'])
        if nm_col and cl_col:
            host_to_cluster = pd.Series(df_h[cl_col].values, index=df_h[nm_col]).to_dict()
            
        df_h = apply_filter(df_h, "Cluster", None)
        
        h_cpu_col = get_col(df_h, ['# CPU', 'vHostNumCPU'])
        h_core_col = get_col(df_h, ['Cores per CPU', 'vHostCoresPerCPU'])
        h_mem_col = get_col(df_h, ['# Memory', 'vHostMemory'])
        
        if h_cpu_col and h_core_col:
            db['cur_cores'] = (to_float(df_h[h_cpu_col]) * to_float(df_h[h_core_col])).sum()
            if h_mem_col: db['cur_total_ram_gb'] = get_rvtools_tb(df_h, h_mem_col) * 1024
            for _, row in df_h.iterrows():
                db['cur_lic_cores'] += (float(row[h_cpu_col]) * max(float(row[h_core_col]), 16))

    elif source_type == "LiveOptics" and 'ESX Hosts' in sheets:
        df_h = promote_header(sheets['ESX Hosts'], ["Host Name", "CPU Cores"])
        
        nm_col = get_col(df_h, 'Host Name')
        cl_col = get_col(df_h, 'Cluster')
        if nm_col and cl_col:
            host_to_cluster = pd.Series(df_h[cl_col].values, index=df_h[nm_col]).to_dict()
            
        df_h = apply_filter(df_h, "Cluster", None)
        
        h_core_col = get_col(df_h, 'CPU Cores')
        h_mem_col = get_col(df_h, 'Memory')
        if h_core_col: db['cur_cores'] = safe_sum(df_h, h_core_col)
        if h_mem_col: db['cur_total_ram_gb'] = to_float(df_h[h_mem_col]).sum() / 1024 / 1024
        
        s_col = get_col(df_h, 'Socket')
        c_col = get_col(df_h, 'Cores')
        if s_col and c_col:
            for _, row in df_h.iterrows():
                try:
                    s = float(row[s_col])
                    c = float(row[c_col])
                    if s > 0: db['cur_lic_cores'] += (s * max(c/s, 16))
                except: pass
    
    if df_h is not None:
        db['cur_host_count'] = len(df_h)

    # 3. Storage Processing (Datastores & vSAN)
    if source_type == "RVTools" and 'vDatastore' in sheets:
        df_ds = promote_header(sheets['vDatastore'], ["Capacity", "vDatastoreCapacity", "Name", "vDatastoreName"])
        
        cap_col = get_col(df_ds, ['Capacity MiB', 'vDatastoreCapacity', 'Capacity'])
        free_col = get_col(df_ds, ['Free MiB', 'vDatastoreFreeSpace', 'Free'])
        type_col = get_col(df_ds, ['Type', 'vDatastoreType'])
        name_col = get_col(df_ds, ['Name', 'vDatastoreName'])
        
        # Map Datastores to Clusters via the "Hosts" list column
        if selected_clusters and "All Clusters" not in selected_clusters and host_to_cluster:
            ds_h_col = get_col(df_ds, ['Hosts', 'vDatastoreHosts'])
            if ds_h_col:
                def map_cluster(hosts_str):
                    if pd.isna(hosts_str): return "Unclustered"
                    for h in str(hosts_str).split(','):
                        h = h.strip()
                        if h in host_to_cluster: return host_to_cluster[h]
                        h_short = h.split('.')[0]
                        for vh in host_to_cluster:
                            if vh.startswith(h_short): return host_to_cluster[vh]
                    return "Unclustered"
                
                df_ds['ClusterMap'] = df_ds[ds_h_col].apply(map_cluster)
                df_ds = df_ds[df_ds['ClusterMap'].isin(selected_clusters)]
                
        mask_vsan = pd.Series([False]*len(df_ds), index=df_ds.index)
        if type_col:
            mask_vsan = mask_vsan | (df_ds[type_col].astype(str).str.lower() == 'vsan')
        if name_col:
            mask_vsan = mask_vsan | (df_ds[name_col].astype(str).str.contains('vsan|vxrail', case=False, na=False))
            
        vsan_ds = df_ds[mask_vsan]
        if not vsan_ds.empty:
            db['vsan_detected'] = True
            db['vsan_raw_tib'] = get_rvtools_tb(vsan_ds, cap_col)
            
        df_san = df_ds[~mask_vsan]
        db['ds_cap'] = get_rvtools_tb(df_san, cap_col)
        db['ds_free'] = get_rvtools_tb(df_san, free_col)
        db['ds_used'] = db['ds_cap'] - db['ds_free']

    elif source_type == "LiveOptics":
        # Live Optics Disk Processing (vSAN Raw)
        disk_tab = next((k for k in sheets.keys() if 'disks' in k.lower() and 'vm' not in k.lower()), None)
        if disk_tab:
            try:
                df_d = promote_header(sheets[disk_tab], ["Model", "Capacity"])
                if selected_clusters and "All Clusters" not in selected_clusters and host_to_cluster:
                    srv_col = get_col(df_d, ['Server', 'Host'])
                    if srv_col:
                        def map_lo_cluster(h):
                            if pd.isna(h): return "Unclustered"
                            h = str(h).strip()
                            if h in host_to_cluster: return host_to_cluster[h]
                            h_short = h.split('.')[0]
                            for vh in host_to_cluster:
                                if vh.startswith(h_short): return host_to_cluster[vh]
                            return "Unclustered"
                        df_d['ClusterMap'] = df_d[srv_col].apply(map_lo_cluster)
                        df_d = df_d[df_d['ClusterMap'].isin(selected_clusters)]
                
                c_mod = get_col(df_d, 'Model')
                c_cap = get_col(df_d, 'Capacity')
                
                if c_mod and c_cap:
                    bad_keywords = 'BOSS|USB|SD|RAID|PERC|Virtual|DVD|CD-ROM|DELL Disk|Cisco Disk'
                    mask_logical = df_d[c_mod].astype(str).str.contains(bad_keywords, case=False, na=False)
                    df_d['TiB'] = to_float(df_d[c_cap]) / 1048576
                    mask_small = df_d['TiB'] < 0.3 
                    
                    df_clean = df_d[~mask_logical & ~mask_small].copy()
                    if not df_clean.empty:
                        grp = df_clean.groupby([c_mod]).agg({'TiB': 'sum', c_mod: 'count'}).rename(columns={c_mod: 'Count'})
                        winner = grp.loc[grp['TiB'].idxmax()]
                        
                        avg_disks = winner['Count'] / db['cur_host_count'] if db['cur_host_count'] else 0
                        if avg_disks > 1.0:
                            db['vsan_detected'] = True
                            db['vsan_raw_tib'] = winner['TiB']
            except: pass

        # Live Optics LUNs
        if 'Host Devices' in sheets:
            df_dev = promote_header(sheets['Host Devices'], ["Device Name", "Capacity"])
            if selected_clusters and "All Clusters" not in selected_clusters and host_to_cluster:
                srv_col = get_col(df_dev, ['Server', 'Host'])
                if srv_col:
                    df_dev['ClusterMap'] = df_dev[srv_col].map(host_to_cluster)
                    df_dev = df_dev[df_dev['ClusterMap'].isin(selected_clusters)]

            type_col = get_col(df_dev, 'Type')
            if type_col:
                df_shared = df_dev[df_dev[type_col].astype(str).str.contains('Cluster', case=False, na=False)]
                cap_col = get_col(df_shared, 'Capacity')
                used_col = get_col(df_shared, 'Used')
                
                if cap_col: db['ds_cap'] = to_float(df_shared[cap_col]).sum() / 1024
                if used_col: db['ds_used'] = to_float(df_shared[used_col]).sum() / 1024 
                db['ds_free'] = db['ds_cap'] - db['ds_used']

    return db

# --- MAIN APP ---
st.sidebar.title("⚙️ Parameters")

# 1. CLUSTER SLOT (Placeholder for dynamic content)
cluster_slot = st.sidebar.empty()

# 2. HARDWARE
st.sidebar.subheader("Target Hardware")
tgt_sockets = st.sidebar.number_input("Sockets", 1, 4, 2)
tgt_cores = st.sidebar.number_input("Cores", 4, 128, 24)
tgt_ram = st.sidebar.number_input("RAM (GB)", 64, 4096, 1024)
host_cap_cores = tgt_sockets * tgt_cores

# 3. CONSTRAINTS
st.sidebar.divider()
vcpu_ratio = st.sidebar.slider("Max vCPU:pCPU", 1.0, 10.0, 5.0, 0.5)
cpu_buffer = st.sidebar.slider("CPU Buffer %", 0, 50, 10)
ram_buffer = st.sidebar.slider("RAM Buffer %", 0, 50, 10)
min_hosts = st.sidebar.number_input("Min Hosts", 1, 32, 3)
ha_nodes = st.sidebar.number_input("HA Nodes", 0, 4, 1)

# 4. SETTINGS
st.sidebar.divider()
include_off = st.sidebar.checkbox("Include Powered Off", True)
growth = st.sidebar.number_input("Growth %", 0.0, 100.0, 10.0) / 100
years = st.sidebar.number_input("Years", 1, 10, 3)
st.sidebar.divider()
cust_name = st.sidebar.text_input("Customer", "Client")
logo_url = st.sidebar.text_input("Logo URL", DEFAULT_LOGO)

# MAIN BODY FILE UPLOAD
st.title(f"📊 {APP_TITLE}")
uploaded_files = st.file_uploader("Upload .xlsx Files", type=["xlsx"], accept_multiple_files=True)

if uploaded_files:
    try:
        file_map = {f.name: f for f in uploaded_files}
        sel_file = st.selectbox("Select File:", list(file_map.keys()))
        
        # Load Raw
        sheets = pd.read_excel(file_map[sel_file], sheet_name=None, header=None, engine='openpyxl')
        sheets = clean_sheet_names(sheets)
        
        # Detect
        src_type = "Unknown"
        df_clus = None
        if 'vInfo' in sheets:
            src_type = "RVTools"
            df_clus = promote_header(sheets['vInfo'], ["VM", "vInfoName"])
        elif 'VMs' in sheets:
            src_type = "LiveOptics"
            df_clus = promote_header(sheets['VMs'], ["VM Name", "Guest Hostname"])
        else:
            st.error("Invalid File Format")
            st.stop()
            
        st.success(f"📂 **{src_type}** Detected")
        
        # --- CHECKBOX LIST SELECTION ---
        clus_col = get_col(df_clus, ["Cluster", "vInfoCluster"])
        selected_clusters = []
        
        with cluster_slot.container():
            st.subheader("Cluster Scope")
            
            if clus_col:
                all_clusters = sorted(df_clus[clus_col].fillna("Unclustered").astype(str).unique())
                
                if not all_clusters:
                    st.caption("No clusters found. Using global scope.")
                    selected_clusters = ["All Clusters"]
                else:
                    st.caption("Check clusters to consolidate:")
                    for c in all_clusters:
                        if st.checkbox(f"{c}", value=True, key=f"chk_{c}"):
                            selected_clusters.append(c)
            else:
                selected_clusters = ["All Clusters"]
                st.caption("No Cluster column found. Using global scope.")

        if not selected_clusters:
            st.warning("⚠️ Please select at least one cluster in the sidebar.")
            st.stop()
            
        # Process
        db = process_data(sheets, src_type, selected_clusters, include_off, "95th")
        
        # Sizing Math
        eff_cores = host_cap_cores * (1 - cpu_buffer/100)
        eff_ram = tgt_ram * (1 - ram_buffer/100)
        
        req_cpu = math.ceil(db['tot_vcpu'] / vcpu_ratio / eff_cores)
        req_ram = math.ceil(db['tot_ram'] / eff_ram)
        
        if req_cpu > req_ram:
            constraint = "CPU"
            raw_hosts = req_cpu
        else:
            constraint = "RAM"
            raw_hosts = req_ram
        
        mult = (1+growth)**years
        fut_hosts = math.ceil(max((db['tot_vcpu']*mult)/vcpu_ratio/eff_cores, (db['tot_ram']*mult)/eff_ram))
        
        hosts_now = max(int(raw_hosts) + ha_nodes, min_hosts)
        hosts_fut = max(int(fut_hosts) + ha_nodes, min_hosts)
        
        # Licensing Calc
        lic_per_node = calc_license_cores(tgt_sockets, tgt_cores)
        now_lic = hosts_now * lic_per_node
        fut_lic = hosts_fut * lic_per_node
        
        # Report Pkg
        rpt = db.copy()
        rpt.update({
            'hosts_now': hosts_now, 'hosts_fut': hosts_fut, 'constraint': constraint,
            'raw_hosts': raw_hosts, 'ha_nodes': ha_nodes, 'sockets': tgt_sockets, 'cores': tgt_cores, 'ram': tgt_ram,
            'tgt_numa_cores': tgt_cores, 'tgt_numa_ram': tgt_ram/tgt_sockets, 'growth': growth, 'years': years,
            'ratio_now': db['tot_vcpu']/(hosts_now*host_cap_cores) if hosts_now else 0,
            'cur_ratio': db['tot_vcpu']/db['cur_cores'] if db['cur_cores'] > 0 else 0,
            'lo_basis': "95th", 'host_cap_cores': host_cap_cores, 'fut_vcpu': db['tot_vcpu']*mult,
            'now_lic_cores': now_lic, 'fut_lic_cores': fut_lic, 'lic_diff': fut_lic - db['cur_lic_cores']
        })
        
        scope_str = ", ".join(selected_clusters) if len(selected_clusters) < 4 else f"{len(selected_clusters)} Clusters Selected"
        
        # DOWNLOAD BUTTON IN SIDEBAR
        safe_cust_name = "".join(c for c in cust_name if c.isalnum() or c in (' ', '-', '_')).strip()
        out_filename = f"{safe_cust_name} - Sizing Report.html"
        
        html = generate_html_report(rpt, scope_str, file_map[sel_file].name, cust_name, logo_url)
        st.sidebar.download_button("Download Full Report", html, file_name=out_filename)

        # Dashboard
        t1, t2 = st.tabs(["📋 Executive Report", "🔍 Raw Data Analysis"])
        with t1:
            st.subheader("1. Executive Sizing Recommendation")
            c1, c2 = st.columns(2)
            with c1:
                st.info(f"### 📅 Current Refresh Requirement")
                st.markdown(f"**Cluster Config:** {hosts_now} Nodes (N+{ha_nodes} HA)")
                
                st.markdown("#### 🔧 Hardware Specs")
                st.write(f"**Per Node:** {tgt_sockets} Sockets | {host_cap_cores} Cores | {tgt_ram} GB RAM")
                st.write(f"**Cluster Total:** {hosts_now * host_cap_cores} Cores | {hosts_now * tgt_ram:,.0f} GB RAM")
                st.markdown("#### 📉 CPU Oversubscription")
                st.metric("Efficiency (vCPU:pCPU)", f"{rpt['ratio_now']:.1f}:1", help="Total Cluster Efficiency (including HA nodes)")
                
                st.divider()
                if constraint == "RAM": st.warning(f"⚠️ **Constraint: Memory Bound**")
                else: st.success(f"✅ **Constraint: CPU Bound**")
                
                with st.expander("View Sizing Logic"):
                    st.write(f"**1. Workload:** {db['tot_vcpu']:,.0f} vCPU, {db['tot_ram']:,.0f} GB RAM")
                    st.write(f"**2. Effective Host:** {eff_cores:.1f} Cores, {eff_ram:.1f} GB RAM")
                    st.write(f"**3. Hosts Needed:** CPU: {req_cpu}, RAM: {req_ram}")
                    st.write(f"**4. Constraint:** {constraint} -> {int(raw_hosts)} active nodes")
                    st.write(f"**5. Final:** {int(raw_hosts)} + {ha_nodes} HA = {hosts_now} Hosts")

            with c2:
                st.success(f"### 🚀 Future Requirement with Growth")
                st.markdown(f"**Cluster Config:** {hosts_fut} Nodes (+{growth*100:.0f}% Growth / {years} Years)")
                st.markdown("#### 🔧 Hardware Specs")
                st.write(f"**Per Node:** {tgt_sockets} Sockets | {host_cap_cores} Cores | {tgt_ram} GB RAM")
                st.write(f"**Cluster Total:** {hosts_fut * host_cap_cores} Cores | {hosts_fut * tgt_ram:,.0f} GB RAM")
            
            st.subheader(f"2. Workload Scope")
            st.caption(f"Consolidating: {scope_str}")
            with st.container(border=True):
                sc1, sc2, sc3, sc4 = st.columns(4)
                sc1.metric("VMs", f"{db['tot_vms']}")
                sc2.metric("vCPU", f"{db['tot_vcpu']:,.0f}")
                sc3.metric("vRAM", f"{db['tot_ram']:,.0f} GB")
                sc4.metric("Current Ratio", f"{rpt['cur_ratio']:.1f}:1")

            st.subheader("3. Storage Requirements")
            c_alloc, c_infra = st.columns(2)
            with c_alloc:
                with st.container(border=True):
                    st.markdown(f"#### 📦 VM Allocation (VMDK)")
                    st.write(f"**Provisioned:** {db.get('vinfo_prov', 0):,.1f} TB")
                    st.write(f"**In Use:** {db.get('vinfo_used', 0):,.1f} TB")
            with c_infra:
                with st.container(border=True):
                    st.markdown(f"#### 🏢 Infrastructure")
                    st.write(f"**Total Capacity:** {db['ds_cap']:,.1f} TB")
                    st.write(f"**Allocated:** {db['ds_used']:,.1f} TB")
                    st.write(f"**Free:** {db['ds_free']:,.1f} TB")
                    
                    if db['vsan_detected']:
                        st.divider()
                        st.markdown(":green[**✅ vSAN/VxRail Detected**]")
                        if db.get('vsan_raw_tib', 0) > 0:
                            st.metric("Raw vSAN Capacity", f"{db['vsan_raw_tib']:,.1f} TiB")
                            st.caption("Note: Raw TiB is used for VVF/VCF Licensing. Usable capacity depends on RAID policy.")
                        else:
                            st.caption("One or more selected clusters contain vSAN-signature disks.")

            st.subheader("4. Architecture & NUMA")
            st.write(f"**Target NUMA:** {tgt_cores} Cores | {tgt_ram/tgt_sockets:.0f} GB RAM")
            
            if db.get('max_vm_cpu', 0) > 0:
                is_wide_cpu = db['max_vm_cpu'] > tgt_cores
                msg_cpu = f"Largest vCPU VM: **{db['name_max_cpu']}** ({db['max_vm_cpu']} vCPU)"
                if is_wide_cpu: st.error(f"⚠️ {msg_cpu} exceeds NUMA cores!")
                else: st.success(f"✅ {msg_cpu} fits NUMA cores.")
                
            if db.get('max_vm_ram', 0) > 0:
                is_wide_ram = db['max_vm_ram'] > (tgt_ram/tgt_sockets)
                msg_ram = f"Largest RAM VM: **{db['name_max_ram']}** ({db['max_vm_ram']:.0f} GB)"
                if is_wide_ram: st.error(f"⚠️ {msg_ram} exceeds NUMA RAM!")
                else: st.success(f"✅ {msg_ram} fits NUMA RAM.")

            st.subheader("5. Licensing")
            l1, l2, l3 = st.columns(3)
            with l1:
                st.metric("Legacy State", f"{db['cur_lic_cores']:,.0f} Cores", f"{db['cur_host_count']} Hosts")
            with l2:
                st.metric("Current Refresh", f"{now_lic:,.0f} Cores", f"{hosts_now} Hosts")
            with l3:
                st.metric("Future w/ Growth", f"{fut_lic:,.0f} Cores", f"Net: {fut_lic - db['cur_lic_cores']:,.0f}")
            
        with t2:
            st.write("Source Data Preview")
            st.dataframe(db.get('df_raw_vinfo'))

    except Exception as e:
        st.error("⚠️ Error Processing File")
        st.code(traceback.format_exc())