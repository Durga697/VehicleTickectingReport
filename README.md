# Vehicle Ticketing Report Automation

Automated pipeline to download Whizzard EV monthly vehicle ticket exports (Return from Ops, Handover to Ops, Service Request, Inter-Station Transfer, and Vehicle Requisition) and generate consolidated daily open-ticket Excel reports.

## Features
- **Headless Playwright Automation**: Downloads all 5 ticket exports automatically.
- **Excel Report Generation**: Formats open tickets, completed tickets, and newly closed tickets with styled summary KPIs and auto-filtered tables.
- **Scheduled GitHub Action**: Runs daily at 10:00 AM IST (04:30 AM UTC), uploads the report as an artifact, and saves the report to the repository.
- **Manual Trigger**: Can be run anytime via the "Run workflow" button in the GitHub Actions tab.

## GitHub Secrets Setup
In your GitHub repository, go to **Settings > Secrets and variables > Actions** and add:
- `WHIZZARD_USER`: Your Whizzard EV login mobile number
- `WHIZZARD_PASS`: Your Whizzard EV password

*(If secrets are not configured, the script defaults to the built-in configuration).*

## Local Setup & Usage

1. **Install Dependencies**:
   ```bash
   pip install -r requirements.txt
   playwright install chromium
   ```

2. **Generate Daily Report**:
   ```bash
   python daily_open_ticket_report.py
   ```

3. **Generate Report from Existing CSVs (Without re-downloading)**:
   ```bash
   python daily_open_ticket_report.py --report-only
   ```
