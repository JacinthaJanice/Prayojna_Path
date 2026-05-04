# Prayojna_Path
Project lifecycle

drive link:
https://drive.google.com/drive/folders/18EJgL5evH_WGUsI512HfTcAXwmhR2vyx?usp=drive_link

# Project Portal - Setup Guide on Linux VM (Ubuntu)

A complete step-by-step guide to run this **Flask-based Project Management Portal** on a Linux Virtual Machine using VMware Workstation.

---

## 🎯 Target VM Specifications (Recommended)

| Component          | Recommended       | Minimum     |
|--------------------|-------------------|-------------|
| Guest OS           | Ubuntu 24.04 LTS  | Ubuntu 22.04 LTS |
| vCPU               | 4 cores           | 2 cores     |
| RAM                | 8 GB              | 4 GB        |
| Disk Space         | 60 GB             | 40 GB       |
| Network            | NAT (default)     | Bridged (for LAN access) |

---

## Step 1: Prepare Host Machine

1. Enable **Virtualization** (Intel VT-x or AMD-V) in your BIOS/UEFI settings.
2. Install **VMware Workstation Pro** (or Player) on your Windows host.
3. Download the latest Ubuntu ISO from [ubuntu.com](https://ubuntu.com/download/desktop).

---

## Step 2: Create Ubuntu VM in VMware Workstation

1. Open VMware Workstation → **Create a New Virtual Machine**.
2. Select **Typical (recommended)**.
3. Choose **Installer disc image file (iso)** and select the Ubuntu ISO.
4. Set VM name: `project-portal-ubuntu`.
5. Set disk size to **60 GB**.
6. Click **Customize Hardware**:
   - Memory: **8192 MB**
   - Processors: **4**
   - Network Adapter: **NAT**
   - (Optional) Enable **Shared Folders**
7. Click **Finish** and power on the VM.

---

## Step 3: Install Ubuntu

- Choose **Normal installation**.
- Enable **third-party software** for graphics and Wi-Fi.
- Create your username and password.
- Complete installation and reboot.
- Remove the ISO when prompted.

---

## Step 4: Install VMware Guest Tools

```bash
sudo apt update
sudo apt install -y open-vm-tools open-vm-tools-desktop
sudo reboot
