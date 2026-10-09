# 📊 Virtualization Sizing Calculator (v2.4)

**Authors:** Adam Marcum & Gemini
**Version:** 2.4

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
    * Upload one or more **RVTools.xlsx** or **Live Optics.xlsx** files. With several files, pick the active one from the dropdown.

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
* Only the needed tabs are read, and the workbook is cached, so changing sidebar settings is instant.

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
* Downloadable HTML report with the AHEAD logo embedded (works offline and when emailed). Preview it in the **Report Preview** tab. A custom logo can be uploaded in the sidebar.

## 📂 Repository Structure

* `sizing_app.py`: The application.
* `requirements.txt`: Required Python libraries.
* `ahead_logo.png` / `ahead_logo_white.png`: AHEAD logos (light / dark backgrounds) embedded in the app and report.
* `Run_on_windows.bat`: Windows installer + launcher.
* `Run_on_mac.command`: macOS installer + launcher.

## 🛠 Configuration Options (Sidebar)

* **Target Hardware:** Sockets, cores per socket, RAM per host, CPU GHz.
* **Constraints:** Max vCPU:pCPU (default 5:1), CPU/RAM overhead, minimum cluster size (default 2), HA tolerance.
* **Host Count Override:** Force a static host count instead of the calculated sizing; shortfalls are flagged.
* **Live Optics Basis:** Peak, Average, or 95th Percentile.
* **Scope:** One, several (consolidation), or all clusters; include/exclude powered-off VMs.

## 📝 Changelog

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
