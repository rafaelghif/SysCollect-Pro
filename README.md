# System Information Collector (VB.NET)

A lightweight, enterprise-grade Windows desktop tool to collect detailed **hardware**, **software**, and **license** information for auditing and asset management purposes. Compatible with Windows XP through Windows 11. All data is exported to structured `.CSV` files organized by hostname and category.

---

## 🚀 Features

- ✅ **Hardware Audit**:
  - Hostname, CPU, RAM, GPU, Disk Model & Size
  - Storage Type (SSD/HDD), TPM, BitLocker
  - Network Interfaces (MAC, IP, Type)
  - BIOS Mode (UEFI/Legacy), Chassis Type, Virtualization Platform
  - Battery Info (for laptops)
  - Disk Partition Style (MBR/GPT)
  - EFI Partition Detection, Boot Drive

- ✅ **Software Audit**:
  - OS Details (Name, Version, Architecture)
  - Installed Applications
  - Antivirus Status

- ✅ **License Audit**:
  - Windows License Status & Product Key
  - Microsoft Office License (via registry and OSPP.vbs)

- 📂 **Auto-foldered by Hostname**
  - `Hardware`, `Software`, and `License` subfolders
  - Sequential `.CSV` files: `HOSTNAME_000.csv`, `HOSTNAME_001.csv`, etc.

- 🖥️ Compatible with:
  - Windows XP, 7, 8, 10, 11
  - Physical and Virtual Machines

- 📎 Clipboard copy of export folder + prompt to open on completion
- 💡 Background loading animation with progress feedback (optional)

---

## 📁 Output Structure

``` plaintext
C:\System Information
├── 01. DESKTOP-1234
│ ├── Hardware
│ │ └── DESKTOP-1234_000.csv
│ ├── Software
│ │ └── DESKTOP-1234_000.csv
│ └── License
│	└── DESKTOP-1234_000.csv
```

Each `.csv` is structured as:

``` plaintext
Property,Value
CPU Model,Intel Core i7-9700
Total RAM,16 GB
...
```

---

## 🔧 How to Use

1. Clone or download this repository
2. Open in Visual Studio (targeting `.NET Framework 4.8`)
3. Build the project
4. Run the executable:  
   `bin\Release\SystemInfoCollector.exe`
5. Click **"Execute"**  
   ✅ Folder copied to clipboard  
   ✅ Prompt to open folder  
   ✅ Files auto-sequenced and organized

---

## 📦 Requirements

- .NET Framework 4.8
- Admin rights *recommended* for accurate TPM, BitLocker, and licensing info
- `cscript` must be available to extract Office license via `ospp.vbs`

---

## ⚠️ Notes & Limitations

- TPM and BitLocker info may not be available on older systems or non-domain devices
- Office license detection varies by version and install type (Click-to-Run, MSI, etc.)
- Partition style detection uses `diskpart` parsing; some results may vary in legacy systems

---

## 🛡️ License

MIT License — feel free to use, modify, and contribute.

---

## 📞 Contact / Contribution

Want to contribute modules (e.g. BIOS updates, asset tag tracking, cloud sync)?  
Open an issue or pull request on GitHub.

For enterprise support, contact: **rafaelghifari.business@gmail.com**

---
