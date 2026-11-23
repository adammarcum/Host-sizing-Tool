# 📊 Virtualization Sizing Calculator (v3.14)

**Authors:** Adam Marcum & Gemini
**Version:** 3.14

A unified infrastructure sizing tool designed for Solutions Architects. It ingests data from **RVTools** or **Dell Live Optics**, performs the sizing math (N+HA, growth, RAM/CPU constraints, licensing) and generates a self-contained HTML executive summary.

---

## 🚀 Quick Start (For Users)

**No coding knowledge required. Nothing to install by hand.**

1.  **Windows:** download the ZIP and unzip it into a normal folder (e.g. Documents), keeping all the files together. **Mac:** skip this step if you use the one-command install below.

2.  **Start it:**

    * **Mac (easiest): one command.** Open **Terminal** (Cmd+Space, type *Terminal*), paste this line and press Return:

      ```
      curl -fsSL https://raw.githubusercontent.com/adammarcum/Host-sizing-Tool/main/install_mac.sh | bash
      ```

      It installs Host Sizer in `~/HostSizer/app` and starts it, with no macOS security prompts. Run the same line again any time to update. Afterwards you can also double-click `Host Sizing Calculator.app` in `~/HostSizer/app`.

    * **Windows:** double-click `Run_on_windows.bat` in the unzipped folder. If Windows shows *"Windows protected your PC"*, click **More info → Run anyway**. If Python isn't installed, the launcher offers to install it for you (answer **Y**).

    * **Mac from the ZIP instead:** macOS blocks apps downloaded from the internet. On macOS 15 (Sequoia) and later, open the app once, then go to **System Settings → Privacy & Security**, scroll down and click **Open Anyway**. Even then macOS may run the app from a temporary copy and report that `Run_on_mac.command` was not found; use the one-command install above instead.

3.  **Wait for your browser to open.** The first run downloads the libraries (about 100 MB, a few minutes). Later runs start in seconds and work offline. Keep the black/Terminal window open while you use the tool; close it to stop.

4.  **Use the tool:**
    * On **Home**, upload one or more **RVTools.xlsx** or **Live Optics.xlsx** files. Each upload joins the current project, so add all your exports first, then click **Configure clusters**; choose **Start a new project** for a separate one. Projects appear as tiles; click **Open** to work on one, or switch in the sidebar.
    * Review **Current Environment**, tune the sizing on **Recommendation**, set the customer name/logo on **Report**, then click **Download consolidated report** (or pick one cluster under **Single reports**).
    * Click **Save project** to keep the sizing. Saved projects appear on **Home** under *Your project library* and reopen with all settings.

### 🧰 If something goes wrong

| What you see | What to do |
|---|---|
| Windows: *"Could not open requirements file"* or *"files were not found next to this launcher"* | The launcher was run from inside the ZIP (Windows only extracts the one file you double-click). Right-click the ZIP → **Extract All...**, then run `Run_on_windows.bat` from the extracted folder. |
| *"The libraries could not be downloaded"* | You need internet for the first run. Connect (and join VPN if required) and try again. A proxy may be blocking `pypi.org` / `files.pythonhosted.org` - ask your eTech team to allow them. |
| Mac: *"Apple could not verify ... is free of malware"* / *"Not Opened"* | Use the one-command install in step 2 (no prompts). Or: **System Settings → Privacy & Security → Open Anyway**. |
| Mac: *"Run_on_mac.command was not found next to this app"* | macOS is running the downloaded app from a temporary copy. The app has already copied the install command to your clipboard: open Terminal, press Cmd+V, then Return. |
| Mac: *"permission denied"* or nothing happens | Use `Host Sizing Calculator.app`, or open Terminal and run `bash ` then drag `Run_on_mac.command` into the window and press Return. |
| Mac: *"app is damaged"* or won't open after an update | In Terminal run: `codesign --force --deep --sign - ` then drag `Host Sizing Calculator.app` into the window and press Return. |
| Windows: Python install declined or `winget` missing | The launcher opens python.org. Install Python 3.12+, tick **Add python.exe to PATH**, then double-click the launcher again. |
| *"Could not create the Python environment"* | The folder is read-only. Move the tool folder to Documents and try again. |
| The tool behaves oddly after an update | Delete the hidden `.venv` folder in the tool folder; the next launch rebuilds it. |

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
* **Storage:** allocated and in-use capacity for VMs, snapshots, templates, RDMs and other datastore files, guest OS used, VM swap, shared datastore capacity, and the effective capacity required (see *Storage consumption* in the 3.13 notes). All storage figures are decimal MB / TB. Live Optics LUNs are de-duplicated across hosts where a device ID is available.

### 4. Licensing
* Current vs. future core licensing with the 16-core-per-CPU minimum. Edition comes from the RVTools vLicense tab or Live Optics ESX Licenses tab; ESXi version is shown separately.

### 5. Professional Reporting
* Downloadable HTML report with the AHEAD logo embedded (works offline and when emailed). Preview it on the **Report** page, where a customer logo can also be uploaded.

* **PowerPoint instead of HTML:** the *Host Sizer report to PowerPoint* Glean skill converts the single or consolidated HTML report into AHEAD-branded slides. Download it from the [shared folder](https://thinkaheadit-my.sharepoint.com/:f:/g/personal/adam_marcum_ahead_com/IgDNWcoha-wgSrplF3-tJ3g4AYkkNnGxvsdONp3Q7LS2zFE?e=b1Gr6H) and add it to Glean.

## 📂 Repository Structure

* `sizing_app.py`: The application.
* `requirements.txt`: Required Python libraries.
* `ahead_logo.png` / `ahead_logo_white.png`: AHEAD logos (light / dark backgrounds) embedded in the app and report.
* `host_sizer_logo_white.png` / `host_sizer_icon.png`: Host Sizer logo for the sidebar and browser tab.
* `Run_on_windows.bat`: Windows installer + launcher.
* `Run_on_mac.command`: macOS installer + launcher.
* `Host Sizing Calculator.app`: Mac double-click shortcut that runs `Run_on_mac.command`.
* `install_mac.sh`: Mac one-command installer / updater (installs to `~/HostSizer/app`).

## 🛠 Configuration Options (Recommendation page → Sizing parameters)

* **Target Hardware:** Processor (optional), sockets, cores per socket, RAM per host (pick a common size or type your own), CPU GHz (base clock).
* **Hypervisor:** Target platform. Sets the Max vCPU:pCPU design limit to its guidance (adjustable) and drives the red ratio flags.
* **Constraints:** CPU/RAM overhead, minimum cluster size (default 2), HA tolerance.
* **Host Count Override:** Force a static host count instead of the calculated sizing; shortfalls are flagged.
* **Live Optics Basis:** Peak, Average, or 95th Percentile.
* **Scope:** One, several (consolidation), or all clusters; size all VMs, powered-on only, or powered-off only (DR).

## 📝 Changelog

**3.14**
* **Recommended host configuration.** For each cluster selection the app picks the best current-generation host (Intel Xeon 6; AMD EPYC 9005 is also considered only when the current hosts already run AMD; otherwise choose AMD manually in the Processor list). **Dual-socket** by default; single-socket only for small clusters (16 cores or fewer per host and 3 hosts or fewer). RAM uses only **balanced memory layouts** (from the Cisco UCS Intel M8 Memory Guide for Xeon 6): every CPU populated identically, all 8 channels (12 on Xeon 6900P / EPYC) filled with 1 or 2 DIMMs per channel, current DIMM sizes 16-256 GB (16 GB 1DPC only; end-of-sale 48 GB left out). On a 2-socket Xeon 6 that means 256, 512, 1,024, 1,536, 2,048, 3,072 or 4,096 GB (no 768 GB), and the recommendation shows the DIMM layout, e.g. *16 x 64 GB DIMMs (1 per channel)*. Dell PowerEdge (R670/R770) and HPE ProLiant Gen12 use the same Xeon 6 memory design (8 channels, 1-2 DIMMs per channel, 16-256 GB DDR5 RDIMMs), so Cisco's rules are the guideline. You can always enter any RAM size manually; Target hardware then shows **Balanced** (with the DIMM layout), a yellow **Unbalanced memory** note for layouts Cisco and Dell support but that leave channels unevenly populated (e.g. 768 GB on 2 sockets = 24 x 32 GB, 12 per CPU: lower bandwidth, not recommended), or an orange **Memory layout not supported** warning with the balanced sizes for that CPU. The Recommendation page's Target hardware card shows the memory layout status, using the same design limit, overhead, HA, minimum cluster and measured CPU as the sizing. Priorities, in order:
    1. **Stay within today's licensed cores.** If no configuration can, it allows up to 10% more than the fewest licensed cores possible, so the host count stays sensible, and says how many cores would be added.
    2. **The largest VM fits one socket** (NUMA), when possible.
    3. **Fewest Day 1 hosts.**
    4. **Least total RAM**, because memory is most of the server cost. Then fewest licensed cores, the most cores for that licence, and the highest clock.
* A new cluster selection now **starts on the recommended configuration** (instead of copying the last one used). Target hardware shows a **Recommended configuration** badge with the reasoning; change any CPU, socket or RAM setting to override (*Custom configuration*), and **Use recommended** restores it. Saved selections and projects keep their settings. The report's Target node card shows *Recommended configuration* when it's in use. Last-generation CPUs stay available in the Processor list.

**3.13**
* **Simpler storage reporting.** The Inventory detail card is now *Datastores* (capacity, used, free, vSAN yes/no, RDM count, scope) and no longer repeats VM figures, which are only under *Storage consumption*. The consumption table shows only the rows that apply (snapshots, templates, RDMs and other files appear when present); snapshot columns and the *without snapshots* totals appear only when the export has snapshots. Scope notes and the vSAN summary use plain wording.
* **ESXi boot and local storage excluded.** Datastore capacity leaves out ESXi boot, service and local datastores (single-host datastores, or names such as *datastore1*, *-service-datastore1*, *local*, *boot*, *scratch*) unless a VM or template lives on them, so stand-alone hosts keep their VM storage. Live Optics *Local* devices without VMs are excluded the same way. The number excluded is shown in the storage scope note. VM and snapshot usage is unaffected (boot volumes aren't VM disks).
* Live Optics host CPU speed is now read from *CPU Clock Speed (GHz)* (newer exports), so current GHz and measured CPU utilization show instead of n/a.
* **Raw device mappings (RDMs).** Current Environment and every report have a *Raw Device Mappings* section for migration planning: each RDM disk with its VM, size, compatibility mode (physical or virtual), sharing (multi-writer or SCSI bus sharing, e.g. Windows failover clusters) and host, plus the LUNs presented to hosts that aren't datastores (RVTools vMultiPath), matched to the RDMs that use them where possible. RDMs are flagged in Risks & flags, and the consolidated Summary has a per-cluster RDM and vSAN table. Sources: RVTools vDisk and vMultiPath (Live Optics exports don't include RDM details).
* **vSAN ESA host storage estimate (optional).** Turn on **Show vSAN options** under Sizing parameters > Storage (off by default). With VMware as the target, the Recommendation page and each report then have a *vSAN Host Storage* section, and the consolidated Summary a per-cluster vSAN table.
    * For the Day 1 host count it finds the capacity drives per host needed to hold the **future** storage need (in use with snapshots, with growth). It shows a row for each common NVMe size (1.6, 1.92, 3.2, 3.84, 6.4, 7.68, 12.8 and 15.36 TB), plus drives for the Day 1 need and for the future host count. It highlights the **closest fit** (least raw capacity) and the **fewest drives**. The options table is for you only: choose **Configuration for the report** (default: closest fit), and the customer report and consolidated Summary show only that configuration and its Broadcom-style capacity breakdown. The choice is kept by drive size, so it follows changes in growth or host count.
    * vSAN ESA only (OSA isn't modelled for new clusters). Protection by cluster size: 2 hosts RAID-1 mirror, 3-5 hosts RAID-5 2+1, 6+ hosts RAID-6 4+2, or choose RAID-5 4+1 (6+ hosts) or another ESA policy. Operations Reserve (Broadcom's ESA formula) from 3 hosts and Host Rebuild Reserve (one host's share, N+1) from 4 hosts, with reserves capped at 30%. Because the reserves provide vSAN's slack, the storage free-space target isn't added (except on 2-host clusters, which have no reserves).
    * The model reproduces Broadcom's vSAN ReadyNode Sizer example (6 hosts x 12 x 1.6 TB, RAID-6, compression 1.2: 51.87 TB). Compression defaults to 1.0 (no data-reduction credit). Confirm final designs in the sizer.
    * **vSAN detection (always on):** RVTools: vSAN datastores (datastore type). Live Optics (no datastore type): a shared datastore whose vendor/model matches the hosts' local disks, or VxRail hosts; it also reports the drive layout per host (e.g. 4 x 7.68 TB + 2 x 0.8 TB cache, vSAN OSA) and how datastore used compares with the VMs' own usage (about 2x suggests RAID-1). When vSAN is found, Current Environment shows a *vSAN detected* badge, the storage details list the vSAN datastores and VMs, and the reports flag it in Risks & flags and the Summary roll-up. For RVTools, VM and template usage on vSAN includes the protection copies and is reduced using today's policy, chosen under Sizing parameters > Storage > *Current environment: vSAN policy today* (shown only when an RVTools source has vSAN; default RAID-1 mirror FTT=1, 2x; OSA and ESA policies are listed because the current cluster may be OSA; *No adjustment* keeps the usage as reported). Live Optics VM usage already excludes the copies, so nothing is removed.
    * New **Storage** tab in Sizing parameters: free space, old snapshot age (moved from Scope & growth), the current vSAN policy, the **Show vSAN options** switch and, when it's on, the vSAN protection policy, compression ratio and max drives per host.
* **Storage consumption.** Current Environment (and every report) has a new *Storage consumption* section for scoping a new array or vSAN:
    * **Allocated** (provisioned) and **In use** for VM disks and files, **snapshots** (count, size, VMs, and how many are older than an age limit, default 30 days), **templates**, **RDMs** (full LUN size) and **other datastore files** (ISOs, orphaned and non-VM files: datastore used minus every VM's and template's usage on its home datastore; datastores shared by several clusters are split between them).
    * Totals **with** and **without snapshots**, plus guest OS used, VM swap files (memory minus reservation, powered-on VMs) and datastore capacity / used / free.
    * **Effective capacity required** for Day 1 and the future, on an in-use and an allocated basis, with all snapshots, without old snapshots and without snapshots: consumed x growth / (1 - free space). Set **Storage free space** (default 20%) and **Old snapshot age** under Sizing parameters > Storage. No vSAN, RAID or data-reduction factors are applied.
    * Sources: RVTools vInfo, vSnapshot, vMemory, vDisk, vPartition, vDatastore and vMetaData (snapshot ages are measured from the export date); Live Optics VMs and its snapshot tab when the export has one. The consolidated Summary has an estate storage roll-up and a per-cluster table.
    * All storage figures are now **decimal MB and TB** (1 TB = 1,000,000 MB); the exports' binary MiB/GiB are converted, so TB values are about 10% higher than before.
    * Saved projects now keep the snapshot, memory and VM path columns. Projects saved earlier still open, but re-add the export to see snapshots, swap and other datastore files.
* **Stand-alone hosts are now sized.** ESXi hosts that aren't in a cluster (and their VMs) were left out because the cluster list only included named clusters; they only counted in *All Clusters*. Each export's unclustered hosts now appear as a **Standalone hosts** entry in Cluster scope ("<source> · Standalone hosts" when a project has several files), so you can size them on their own, combine them with clusters, and they appear in the reports. Projects saved earlier pick this up when opened.
* **Mac launcher works from a downloaded ZIP.** macOS runs apps from a downloaded (quarantined) ZIP from a temporary copy, so *Host Sizing Calculator* couldn't find *Run_on_mac.command* next to it. It now finds the real Host Sizer folder (Spotlight, then Downloads, Desktop, Documents, ~/HostSizer and OneDrive/iCloud folders), starts from there and clears the download's quarantine flag so later starts are normal. If the folder can't be found, it copies the one-command install to the clipboard and explains what to do.

**3.12**
* Launchers check that the tool files are next to them and explain how to fix it (usually: extract the ZIP first) instead of failing with *"Could not open requirements file"*.
* **Combine RVTools with Live Optics.** When a project has both, Home shows a pairing table: for each RVTools cluster, choose the Live Optics capture of the same cluster. Paired clusters take VMs, vCPU and RAM from RVTools and measured CPU (GHz) from Live Optics, and the Live Optics cluster isn't counted separately. Each pair shows a confidence (*High*, *Possible*, *Low*, or *Conflict*) from host, core, RAM and VM counts and shared VM/host names. *Conflict* flags clusters whose (often anonymised) names match but whose sizes don't. **Fill in suggestions** pre-selects likely matches; nothing is paired until you click **Save pairs**. Pairs are saved with the project and noted in the reports.
* **Workload labels.** Select a cluster to give it a label (e.g. SQL Server, Citrix, DMZ). A suggestion from the cluster or VM names is offered for you to accept or ignore. Labels appear in the page header, the report heading and the consolidated cluster comparison, and are saved with the project.
* **Measured CPU (GHz) sizing.** With Live Optics data, hosts are now sized three ways: CPU by vCPU ratio, CPU by measured GHz (the chosen basis at an 80% target, after overhead), and RAM. The largest wins and the binding one is named. GHz is also checked in the Day 1 design check (with the HA hosts failed) and is a third limit in Room to grow.
* **Processor picker.** Target hardware has a **Processor** list of current and previous-generation server CPUs (Intel Xeon 6 and 5th Gen Xeon Scalable; AMD EPYC 9005 Turin and 9004 Genoa), grouped by generation and sorted by core count. Type to search (e.g. *6737P*). Picking one fills Cores / socket and the base clock (GHz); 1-socket-only parts set Sockets to 1. Editing cores or GHz by hand switches back to *Custom*. The processor name appears in the reports and is saved per cluster selection with the project. This replaces the Hardware options comparison and catalog.
* **RAM per host dropdown.** Choose a common configuration (256, 512, 768, 1,024, 1,536, 2,048, 3,072 or 4,096 GB), or type another amount (64-8,192 GB, e.g. *1280* or *1,280 GB*) and press Enter. Amounts that aren't a multiple of 256 GB get a warning to check the DIMM layout.
* The version number now appears only under **About** in the sidebar (no longer under the logo), and the *Update available* notice shows there too.

**3.11**
* The app version is shown under the Host Sizer logo in the sidebar.
* **Remove a sizing.** Next to the *Configured* list under Cluster scope, **Manage** lists every saved cluster selection with an **x** to remove it (for example a multi-cluster combination changed by accident). Removed sizings drop out of the reports; selecting the same clusters again starts fresh.
* **Sizing method in every report.** A short *Sizing Method* section explains, in plain language, how the sizing was done (data, scope, usable host, host count, checks, room to grow) and lists the settings used for each cluster. It appears on the Summary tab of the consolidated report and at the end of each single report.
* **VMs to size: All VMs, Powered on only, or Powered off only.** The new three-way choice (Sizing parameters → Scope & growth) replaces the powered-off toggle. *Powered off only* sizes a DR environment; suspended VMs count as powered off, and Live Optics performance sizing is hidden in that mode. Existing projects keep their previous choice.
* **Room to grow.** Each sizing shows the average VM (vCPU and RAM) and how many more average VMs fit on the Day 1 and future clusters, with the HA hosts failed and the design limit, before another host is needed, plus how many each extra host adds. It's on the Recommendation page, in each report, and as a per-cluster table on the consolidated Summary. (Compute only; storage is under Storage consumption.)

**3.10**
* **Consolidated report.** *Download consolidated report* (the default) creates one HTML file, `<customer>-Consolidated-AHEAD Sizing.html`, with a **Summary** tab and a tab for every sized cluster. The Summary shows totals for current vs. Day 1 vs. future, a cluster comparison table (click a name to open its tab), a licensing and storage roll-up, and a list of risks and flags by cluster. Print / Save PDF prints every tab.
* Consolidations (e.g. `C1 + C2`) and an edited *All Clusters* selection appear as **scenario** tabs and aren't added to the totals, so nothing is counted twice.
* **Several exports in one project.** On **Home**, add as many RVTools and/or Live Optics files as you need: each upload joins the current project (choose **Start a new project** to begin a separate one), and you stay on Home until you click **Configure clusters**. Every cluster from every file then appears under **Cluster scope**, ready to click through and size. With two or more sources, clusters are named `source · cluster` (the source name defaults to the vCenter, and can be renamed on Home), so clusters from different vCenters stay separate and can be selected together to size a cross-site consolidation. The consolidated report covers every source.
* Each source can be included or excluded, or removed. Sources that share VM or host names are flagged as a *possible overlap* so the same environment isn't counted twice. Performance sizing is shown when every cluster in a selection comes from Live Optics. Saved projects keep all sources (only the columns the sizing needs); older projects open unchanged.
* **Out-of-scope clusters.** Select a single cluster and turn off **Active** (under Cluster scope) to take it out of scope: it isn't sized, can't be configured, and is left out of All Clusters, consolidations and every report. Out-of-scope clusters are marked in the cluster tags and saved with the project.
* The consolidated report's cluster tabs wrap onto several rows, so every cluster is visible without scrolling sideways.
* Each cluster or scenario can still be downloaded on its own from **Single reports**. The Report page lists what the consolidated report contains, notes any clusters not yet sized, and previews either the consolidated or the single report.

**3.9**
* **Easier Mac install.** New one-command installer and updater: paste `curl -fsSL https://raw.githubusercontent.com/adammarcum/Host-sizing-Tool/main/install_mac.sh | bash` into Terminal. It installs to `~/HostSizer/app` and starts the tool without macOS security prompts, and re-running it updates in place (your environment and projects are kept).
* `Host Sizing Calculator.app` now explains macOS's "temporary copy" problem (App Translocation) instead of only saying `Run_on_mac.command` was not found, and copies the install command to your clipboard.
* Install instructions updated for macOS 15 (Sequoia), where right-click → Open no longer bypasses the security prompt.

**3.8**
* **Update notice.** When a newer version is published on GitHub, an *Update available* card appears under **About** in the sidebar with a **Download update** button and a link to what's new (plus a one-time pop-up). The check reads only the version number from GitHub, at most every 6 hours, and is skipped silently when offline. Set `HOSTSIZER_NO_UPDATE_CHECK=1` to turn it off.

**3.7**
* Downloaded reports are named `<customer>-<cluster>-AHEAD Sizing.html` (e.g. `Acme Corp-Prod-AHEAD Sizing.html`; consolidations read `Acme Corp-C1 + C2-AHEAD Sizing.html`, the full environment `Acme Corp-All Clusters-AHEAD Sizing.html`).

**3.6**
* **Settings per cluster.** Every cluster selection (each single cluster, any consolidation combination, and All Clusters) keeps its own sizing settings: hypervisor, hardware, constraints, override, growth and powered-off handling. A selection you haven't configured yet starts from the one you were just on. Saved projects store all of them, so one project covers every cluster in the export; click through the clusters to review each configuration. The customer name is shared across the project.
* If the export has a single cluster, it is selected automatically instead of "All Clusters".
* The downloadable report covers the cluster selection on screen. Projects saved by older versions open normally.

**3.5**
* Host Sizer logo (draft, pending Brand Marketing review) in the sidebar and as the browser-tab icon. The HTML report is unchanged.

**3.4**
* Easier first run. The Windows launcher finds a real Python (ignoring the Microsoft Store placeholder), checks standard install folders, and offers to install Python 3.12 with winget; otherwise it opens python.org. The Mac launcher prefers a newer python.org/Homebrew Python, falls back to Apple's, and installs the Command Line Tools when needed.
* Both launchers rebuild a damaged `.venv`, explain download failures (internet, VPN, proxy), and skip Streamlit's first-run email prompt. Streamlit usage statistics are turned off.
* `Host Sizing Calculator.app` now runs `Run_on_mac.command`, so it no longer depends on the Mac's built-in Python libraries and works even if the launcher lost its execute permission.

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
