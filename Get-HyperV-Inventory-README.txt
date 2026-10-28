Hyper-V Inventory Collection Instructions
=========================================

This script gathers the hardware specifications (CPU/RAM) of your Hyper-V hosts 
and the allocation details of your Virtual Machines. It runs safely using 
standard Windows management tools (WMI/PowerShell).

Prerequisites:
--------------
1. You must be a Local Administrator on the Hyper-V hosts.
2. If scanning a Cluster, run this from one of the Cluster Nodes or a 
   management server with RSAT tools installed.
3. PowerShell must be allowed to run scripts.

How to Run:
-----------
1. Download the attached file "Get-HyperV-Inventory.ps1" to a folder on the server (e.g., C:\Temp).
2. Right-click the file and select "Properties".
3. If you see a checkbox labeled "Unblock" at the bottom, check it and click OK.
4. Right-click the file and select "Run with PowerShell".
   (Or open PowerShell as Administrator, navigate to the folder, and run: .\Get-HyperV-Inventory.ps1)

Output:
-------
The script will generate two files in the same folder:
1. HyperV_vInfo.csv (VM Data)
2. HyperV_vHost.csv (Host Hardware Data)

Please zip these two files and return them to us for the sizing analysis.