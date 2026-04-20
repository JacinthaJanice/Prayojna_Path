# Prayojna_Path
Project lifecycle

drive link:
https://drive.google.com/drive/folders/18EJgL5evH_WGUsI512HfTcAXwmhR2vyx?usp=drive_link

# Project Portal

Project Portal is a Flask-based web application for managing projects, project details, meetings, costs, analysis, closures, and report exports.

## What You Need

- Windows PC or a Windows virtual machine
- Python 3.10 or newer
- `pip`
- `wkhtmltopdf` for PDF export

## Install on Your PC or VMware VM

1. Create a folder for the project and copy all files into it.
2. If you are using VMware, create a new Windows virtual machine first and copy the project folder into the VM.
3. Open a terminal in the project folder.
4. Create and activate a virtual environment:

```powershell
py -m venv .venv
.venv\Scripts\Activate.ps1
```

4. Install dependencies:

```powershell
pip install -r requirements.txt
```

5. Install `wkhtmltopdf` if you want PDF download support.
6. If `wkhtmltopdf` is not in PATH, set the `WKHTMLTOPDF_PATH` environment variable to the full executable path.
7. If you are using VMware Tools, install them inside the VM for better mouse, display, and clipboard integration.

## Run the Application

You can start the app in either of these ways:

```powershell
python app.py
```

or

```powershell
python start.py
```

The app will open in your browser at:

```text
http://127.0.0.1:5000/
```

## Default Login

If the database is being created for the first time, the app seeds a default admin user:

- Username: `admin`
- Password: `admin@72$`

Change the password after first login.

## Notes for VM Installation

- Make sure the VMware VM has network access if you need to install Python packages.
- If you are sharing the app between host and VM, keep all project files inside the VM or use a shared folder carefully.
- If PDF export fails, verify that `wkhtmltopdf` is installed and reachable from the VM.

## Suggested Screenshots for the README

You can add screenshots in a folder such as `screenshots/` and reference them in the README.

Recommended screenshots:

1. VMware VM setup or Windows desktop inside the VM.
2. Terminal showing virtual environment activation and dependency installation.
3. Browser showing the Project Portal login page.
4. Dashboard or project list page after login.
5. Add Project form.
6. Project details or analysis page.
7. PDF download output or exported report preview.

Example markdown if you want to add images later:

```markdown
![Login Page](screenshots/login-page.png)
![Dashboard](screenshots/dashboard.png)
```

## Main Files

- `app.py` - main Flask application
- `start.py` - launches the app and opens the browser
- `requirements.txt` - Python dependencies
- `templates/` - HTML templates
- `static/` - CSS, JS, and other static files

