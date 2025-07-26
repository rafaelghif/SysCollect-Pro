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
