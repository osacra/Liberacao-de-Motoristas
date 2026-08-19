# Sistema de Liberação de Motoristas

Desktop automation built with Python to monitor operational spreadsheets, identify pending driver releases and notify the responsible team. The project was created from a real logistics workflow and focuses on reducing repetitive manual work, improving validation and preserving an operational log.

## Impact

The system automates a driver-release process in a 24/7 logistics operation. The solution centralizes spreadsheet monitoring, compares new and processed records, identifies pending releases and sends notifications. In the professional context where it was developed, the automation reduced approximately **56 hours of manual effort per month**.

## Main capabilities

- Continuous monitoring of Excel workbooks.
- Comparison between new and already processed records.
- Validation and processing of pending driver releases.
- Automatic e-mail notifications.
- Selenium integration with Microsoft Edge and Excel Online.
- Desktop interface with start, pause, status and history controls.
- Structured logging, alerts and modular business rules.
- Local configuration for Windows-based operational environments.

## Architecture

```text
main.py
  |
  +-- ui/       Desktop interface with PySide6
  +-- core/     Monitoring, processing, Excel, e-mail and alert rules
  +-- config/   Local configuration and environment-specific settings
  +-- assets/   Application and e-mail assets
```

The separation between the UI, processing rules and integration handlers makes it possible to change the operational workflow without coupling every component to the desktop interface.

## Technology stack

| Area | Technologies |
|---|---|
| Language | Python 3 |
| Data and spreadsheets | Pandas, OpenPyXL |
| Automation and integration | Selenium, Requests, Microsoft Edge WebDriver |
| Desktop UI | PySide6, PyStray, Pillow |
| Notifications | Outlook integration through `pywin32` |
| Quality and operations | Modular structure, configuration file and logging |

## Run locally on Windows

This project integrates with Microsoft Excel, Outlook and Edge WebDriver, so it is intended for a Windows environment with those dependencies available.

### Requirements

- Python 3.10 or newer.
- Microsoft Edge and a compatible Edge WebDriver.
- Microsoft Excel and Outlook when using the corresponding integrations.
- Access to the operational workbook and its permitted notification workflow.

### Install

```powershell
git clone https://github.com/osacra/Liberacao-de-Motoristas.git
cd Liberacao-de-Motoristas
python -m venv .venv
.\\.venv\\Scripts\\Activate.ps1
pip install -r requirements.txt
```

### Configure

Copy the example configuration and replace only the local values required by your environment:

```powershell
Copy-Item config\\config.example.json config\\config.json
```

Configure workbook paths, the permitted Excel Online URL, notification recipients, Edge WebDriver path and local asset paths. The file `config/config.json` is ignored by Git and must never contain credentials or private production links in a public commit.

### Start

```powershell
python main.py
```

## Important security and portability notes

The automation depends on local workbooks, Microsoft accounts and Windows integrations. Do not commit private SharePoint links, personal e-mail addresses, access tokens, exported spreadsheets or machine-specific paths. Use the example configuration as a template and keep operational data outside the repository.

The project is an internal-process automation and is not designed to run unchanged in a public cloud environment. A future evolution could replace local Outlook/Edge dependencies with service APIs, environment variables and a controlled job runner.

## Future improvements

1. Add automated tests around spreadsheet parsing and business rules.
2. Replace hardcoded local paths with validated configuration values.
3. Add structured error reporting and retry policies for external integrations.
4. Provide a sanitized demo dataset for reproducible portfolio demonstrations.
5. Package the desktop application with a documented release process.

## Author

Developed by [Arthur Sacramento](https://github.com/osacra).
