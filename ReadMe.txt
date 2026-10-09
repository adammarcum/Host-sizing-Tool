# 📊 Virtualization Sizing Calculator (v3.3)

**Authors:** Adam Marcum & Gemini
**Version:** 3.3

A unified infrastructure sizing tool designed for Solutions Architects. It ingests data from **RVTools** or **Dell Live Optics**, performs the sizing math (N+HA, growth, RAM/CPU constraints, licensing) and generates a self-contained HTML executive summary.

---

## 🚀 Quick Start (For Users)

**No coding knowledge required.**

1.  **Download the Tool:**
    * Click the green **<> Code** button above, select **Download ZIP**, and extract it (e.g., to Documents).
    * Keep all files together in the same folder (`sizing_app.py`, `requirements.txt`, both `ahead_logo` files, launchers).

2.  **Launch:**
    * **Windows:** Double-click `Run_on_windows.bat`.
    * **Mac:** Double-click `Run_on_mac.command`.
        * If macOS says it can't be opened, right-click the file and choose **Open**. If it says "permission denied", run `chmod +x Run_on_mac.command` once in Terminal.

    The first run creates a private Python environment (`.venv`) and installs the libraries; later runs start immediately. Your browser opens automatically. Close the terminal window to stop the app.

3.  **Use:**
    * On **Home**, upload one or more **RVTools.xlsx** or **Live Optics.xlsx** files. Each appears as a tile; click **Open** to work on it, or switch the active report in the sidebar.
    * Review **Current Environment**, tune the sizing on **Recommendation**, set the customer name/logo on **Report**, then click **Download report**.
    * Click **Save project** to keep the sizing. Saved projects appear on **Home** under *Your project library* and reopen with all settings.

---

## 💻 Developer Start (Manual Method)

**Prerequisites:** Python 3.9+

```bash
python3 -m venv .venv
source .venv/bin/activate        # Windows: .venv\Scripts\activate
pip install -r requirements.txt
streamlit run sizing_app.py
```

The sizing math lives in `compute_sizing()` and has no Streamlit dependencies, so it can be unit-tested directly.

---

## ✨ Key Features

### 1. Universal Data Ingestion
* **RVTools Mode:** Sizes on *Entitlement* (allocated vCPU/vRAM). Templates and SRM placeholders are excluded.
* **Live Optics Mode:** Adds *Consumption* sizing (Peak, Average, or 95th Percentile GHz).
* Only the needed tabs are read, and the workbook is cached, so changing sizing parameters is instant.

### 2. Sizing Logic
* **Constraint Analysis:** CPU bound, Memory bound, or Balanced.
* **HA & Growth:** N+X redundancy, minimum cluster size, compound annual growth.
* **Overhead Awareness:** Adjustable hypervisor CPU and RAM buffers (default 10%).
* **Ratios:** vCPU:pCPU shown for all hosts and during an HA failover, against your design limit.
* **Performance Sizing:** GHz-based host count at 80% target utilization, never below the RAM requirement.

### 3. Architecture & Storage
* **NUMA Analysis:** Counts VMs wider than a target socket (CPU and RAM separately) and names the largest.
* **Storage:** VMDK provisioned/in-use, guest OS used (backup scope), RDM disks, and shared datastore capacity. Live Optics LUNs are de-duplicated across hosts where a device ID is available.

### 4. Licensing
* Current vs. future core licensing with the 16-core-per-CPU minimum. Edition comes from the RVTools vLicense tab or Live Optics ESX Licenses tab; ESXi version is shown separately.

### 5. Professional Reporting
* Downloadable HTML report with the AHEAD logo embedded (works offline and when emailed). Preview it on the **Report** page, where a customer logo can also be uploaded.

## 📂 Repository Structure

* `sizing_app.py`: The application.
* `requirements.txt`: Required Python libraries.
* `ahead_logo.png` / `ahead_logo_white.png`: AHEAD logos (light / dark backgrounds) embedded in the app and report.
* `Run_on_windows.bat`: Windows installer + launcher.
* `Run_on_mac.command`: macOS installer + launcher.

## 🛠 Configuration Options (Recommendation page → Sizing parameters)

* **Target Hardware:** Sockets, cores per socket, RAM per host, CPU GHz.
* **Hypervisor:** Target platform. Sets the Max vCPU:pCPU design limit to its guidance (adjustable) and drives the red ratio flags.
* **Constraints:** CPU/RAM overhead, minimum cluster size (default 2), HA tolerance.
* **Host Count Override:** Force a static host count instead of the calculated sizing; shortfalls are flagged.
* **Live Optics Basis:** Peak, Average, or 95th Percentile.
* **Scope:** One, several (consolidation), or all clusters; include/exclude powered-off VMs.

## 📝 Changelog

**3.3**
* **Saved projects.** Use **Save project** at the top of any page to keep a sizing and reopen it later from **Home → Your project library**. A project stores your settings (hypervisor, hardware, constraints, override, growth, cluster scope, customer name and logo) plus only the columns the sizing uses; everything else in the export (IP addresses, annotations, networks, etc.) is discarded.
* Projects are saved on this computer in `~/HostSizer/Projects` as `.hsproj` files (compressed JSON). Set the `HOSTSIZER_PROJECTS` environment variable to use a different folder.
* Each save is checked: the trimmed data must give exactly the same sizing as the full export, or the save is refused.
* Page headers show Saved / Unsaved changes / Not saved. Saving under the same name updates the project; projects can be deleted from Home.

**3.2**
* The Max vCPU:pCPU design limit moved to the Hypervisor tab and is set automatically to the chosen hypervisor's guidance (default VMware 4.5:1). It can still be adjusted, with a one-click reset. Basis column removed from the hypervisor table.

**3.1**
* New **Hypervisor** tab in Sizing parameters (VMware, Nutanix, Azure Local, Hyper-V) with vCPU:pCPU guidance: VMware 4.5:1 (average of Intel-captured workloads), Nutanix 3.5:1 (average), Azure Local 2.5:1 (maximum recommended by Intel and Dell product management), Hyper-V 2:1 (average).
* Any vCPU:pCPU ratio above the selected hypervisor's guidance is flagged in **bold red** in the app and the HTML report (current, Day 1, future; all hosts and HA failover). One-click option to set the design limit to the guidance value.

**3.0**
* New app UI in the AHEAD tool design language (Roost / Hatch): "Host Sizer · Powered by AHEAD" sidebar, dark navy navigation rail with icons, white bordered cards, Poppins type, AHEAD blue buttons, line icons and status badges instead of emoji and banners.
* Pages: **Home** (greeting, upload, "Jump back in" report tiles), **Current Environment**, **Recommendation** (sizing parameters in tabs, recommendation cards, Day 1 design check, Current vs. Recommended with Day 1 / Future / Both toggle, NUMA, licensing, performance), **Report** (customer name and logo, preview), **Raw Data** (VMs / Hosts, key or all columns, search).
* Download report button on every page header; cluster scope chosen with selectable tags.
* AHEAD theme is written to `.streamlit/config.toml` automatically (launchers do this before starting; the Mac .app does it on first run, then asks for a restart). A config file you manage yourself is never overwritten.
* Requires Streamlit 1.50 or newer (the launchers upgrade automatically). HTML report design unchanged.

**2.4**
* HTML report restyled to AHEAD brand standards (2024 system): Poppins/PT Serif, AHEAD palette, navy hero with headline stats, Day 1 design-check panel, card layout, sticky section nav, Print / Save PDF button, confidential footer. Passes the AHEAD brand validator.
* Sidebar logo upload is now the **customer** logo (shown in the 'Prepared for' chip); AHEAD logos are embedded automatically.

**2.3**
* New sidebar **Host Count Override**: force a static node count (optionally for future sizing too). The calculated count is still shown, and the forced count is validated against the vCPU:pCPU limit, usable RAM, HA and minimum cluster size. Licensing and ratios follow the forced count.

**2.2**
* Screen and report now open with **Current Environment** (hosts, cores, RAM, CPU model/speed, vCPU:pCPU, vRAM vs host RAM, N-1 ratios, VM averages, storage, current licensing), followed by the Sizing Recommendation.
* New Current vs. Recommended vs. Future comparison table.

**2.1**
* Fixed `requirements.txt` typo (`streamline` → `streamlit`) and pinned version ranges.
* Report "Design Limit" now shows the configured limit; screen and report use the same ratios.
* NUMA check names the correct VM for CPU vs. RAM and counts all wide VMs.
* Performance sizing now respects the RAM requirement.
* License edition read from vLicense (RVTools); templates/SRM placeholders excluded.
* Live Optics storage scoped to selected clusters and de-duplicated across hosts.
* Cached, selective workbook loading; vectorized licensing math.
* HTML-escaped report content; embedded logo; multi-file upload; clearer error details.
* Launchers use a private virtual environment and launch the app on both platforms.

## 📜 License

This project is open for internal use and modification.
