"""
Enhancement Jobs Report Generator
Pulls data from Jobber API and generates monthly Excel report
Runs weekly on Fridays via GitHub Actions

Supports two modes:
1. API Mode: Fetches data directly from Jobber GraphQL API
2. CSV Mode: Reads from exported CSV files in Downloads folder

CSV fallback is used when API doesn't return expected data or for local testing.
"""

import os
import csv
import glob
import json
import requests
from datetime import datetime, timedelta
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.formatting.rule import FormulaRule, DataBarRule

# Configuration
API_URL = "https://api.getjobber.com/api/graphql"
API_VERSION = "2023-11-15"

# Color palette
NAVY = "1B365D"
DARK_BLUE = "2E5090"
MEDIUM_BLUE = "4472C4"
LIGHT_BLUE = "D6DCE5"
GREEN_LIGHT = "C6EFCE"
RED_LIGHT = "FFC7CE"
YELLOW_LIGHT = "FFEB9C"
ENHANCEMENT_FILL = "E2EFDA"
CONTRACTED_FILL = "DDEBF7"

# Styles
header_fill = PatternFill(start_color=NAVY, end_color=NAVY, fill_type="solid")
header_font = Font(bold=True, color="FFFFFF", size=10, name='Calibri')
title_font = Font(bold=True, size=18, color=NAVY, name='Calibri')
section_font = Font(bold=True, size=12, color=NAVY, name='Calibri')
total_fill = PatternFill(start_color=MEDIUM_BLUE, end_color=MEDIUM_BLUE, fill_type="solid")
total_font = Font(bold=True, size=10, color="FFFFFF", name='Calibri')
green_fill = PatternFill(start_color=GREEN_LIGHT, end_color=GREEN_LIGHT, fill_type="solid")
red_fill = PatternFill(start_color=RED_LIGHT, end_color=RED_LIGHT, fill_type="solid")
alt_row_fill = PatternFill(start_color="F8F9FA", end_color="F8F9FA", fill_type="solid")
thin_border = Border(
    left=Side(style='thin', color=LIGHT_BLUE),
    right=Side(style='thin', color=LIGHT_BLUE),
    top=Side(style='thin', color=LIGHT_BLUE),
    bottom=Side(style='thin', color=LIGHT_BLUE)
)
thick_bottom = Border(bottom=Side(style='medium', color=NAVY))
currency_format = '_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)'
currency_format_whole = '_($* #,##0_);_($* (#,##0);_($* "-"??_);_(@_)'
percent_format = '0.0%'


# CSV file locations
DOWNLOADS_DIR = r"C:\Users\daria\Downloads"
JOBS_CSV_PATTERN = "One-off jobs_Report_*.csv"
INVOICES_CSV_PATTERN = "Invoices_Report_*.csv"


def find_latest_csv(pattern, directory=DOWNLOADS_DIR):
    """Find the most recent CSV file matching pattern"""
    search_path = os.path.join(directory, pattern)
    files = glob.glob(search_path)
    if not files:
        return None
    # Sort by modification time, newest first
    return max(files, key=os.path.getmtime)


def parse_currency(value):
    """Parse currency string to float"""
    if not value:
        return 0.0
    if isinstance(value, (int, float)):
        return float(value)
    # Remove $ and commas, handle parentheses for negatives
    cleaned = str(value).replace('$', '').replace(',', '').strip()
    if cleaned.startswith('(') and cleaned.endswith(')'):
        cleaned = '-' + cleaned[1:-1]
    try:
        return float(cleaned)
    except ValueError:
        return 0.0


def parse_percentage(value):
    """Parse percentage string to float"""
    if not value:
        return 0.0
    cleaned = str(value).replace('%', '').strip()
    try:
        return float(cleaned) / 100.0
    except ValueError:
        return 0.0


def parse_date(date_str):
    """Parse date from various formats"""
    if not date_str or date_str == '-':
        return None
    date_str = date_str.strip()
    # Try multiple date formats
    formats = [
        '%Y-%m-%d',           # 2026-01-31
        '%b %d, %Y',          # Jan 31, 2026
        '%B %d, %Y',          # January 31, 2026
        '%m/%d/%Y',           # 01/31/2026
        '%m/%d/%y',           # 01/31/26
    ]
    for fmt in formats:
        try:
            return datetime.strptime(date_str, fmt).date()
        except ValueError:
            continue
    return None


def load_jobs_from_csv(csv_path, start_date, end_date):
    """Load and filter jobs from CSV export"""
    jobs = []

    with open(csv_path, 'r', encoding='utf-8-sig') as f:
        reader = csv.DictReader(f)
        for row in reader:
            # Check job type - only Enhancement and Contracted Enhancement
            job_type = row.get('Job Type', '').strip()
            if job_type not in ('Enhancement', 'Contracted Enhancement'):
                continue

            # Check date range using Scheduled start date
            date_str = row.get('Scheduled start date', '')
            job_date = parse_date(date_str)
            if job_date:
                if not (start_date <= job_date <= end_date):
                    continue

            # Parse financial data
            revenue = parse_currency(row.get('Total revenue ($)', 0))
            cost = parse_currency(row.get('Total costs ($)', 0))
            profit = parse_currency(row.get('Profit ($)', 0))

            jobs.append({
                'jobNumber': row.get('Job #', ''),
                'title': row.get('Job Type', ''),  # Use Job Type as title since that's what we're filtering on
                'client': {'name': row.get('Client name', '')},
                'jobStatus': 'closed' if row.get('Closed date') else 'active',
                'startAt': row.get('Scheduled start date', ''),
                'closedAt': row.get('Closed date', ''),
                'salesperson': row.get('Salesperson', ''),
                'assignedTo': row.get('Visits assigned to', ''),
                'invoiceNumbers': row.get('Invoice #s', ''),
                'expensesTotal': parse_currency(row.get('Expenses total ($)', 0)),
                'timeTracked': row.get('Time tracked', ''),
                'labourCost': parse_currency(row.get('Labour cost total ($)', 0)),
                'jobCosting': {
                    'totalRevenue': revenue,
                    'totalCost': cost,
                },
                'total': revenue,
                '_profit': profit,
                '_profit_pct': parse_percentage(row.get('Profit %', 0)),
                '_job_type': job_type,
                '_is_enhancement': job_type == 'Enhancement',
                '_is_contracted': job_type == 'Contracted Enhancement',
            })

    return jobs


def load_invoices_from_csv(csv_path, job_numbers):
    """Load and filter invoices from CSV export"""
    invoices = []

    with open(csv_path, 'r', encoding='utf-8-sig') as f:
        reader = csv.DictReader(f)
        for row in reader:
            # Check if invoice is related to any of our job numbers
            invoice_jobs = row.get('Job #s', '')
            related = False
            for jn in job_numbers:
                if str(jn) in invoice_jobs:
                    related = True
                    break

            if not related:
                continue

            invoices.append({
                'invoiceNumber': row.get('Invoice #', ''),
                'client': {'name': row.get('Client name', '')},
                'job': {'jobNumber': invoice_jobs.split(',')[0].strip() if invoice_jobs else ''},
                'createdDate': row.get('Created date', ''),
                'issuedDate': row.get('Issued date', ''),
                'dueDate': row.get('Due date', ''),
                'lateBy': row.get('Late by', ''),
                'paidDate': row.get('Marked paid date', ''),
                'daysToPaid': row.get('Days to paid', ''),
                'lastContacted': row.get('Last contacted', ''),
                'invoiceStatus': row.get('Status', ''),
                'total': parse_currency(row.get('Total ($)', 0)),
                'balance': parse_currency(row.get('Balance ($)', 0)),
            })

    return invoices


def get_access_token():
    """Get access token from environment or refresh if needed"""
    access_token = os.environ.get('JOBBER_ACCESS_TOKEN')
    refresh_token = os.environ.get('JOBBER_REFRESH_TOKEN')
    client_id = os.environ.get('JOBBER_CLIENT_ID')
    client_secret = os.environ.get('JOBBER_CLIENT_SECRET')

    # Try existing token first
    if access_token:
        test_query = '{ jobs(first: 1) { totalCount } }'
        response = requests.post(
            API_URL,
            headers={
                'Authorization': f'Bearer {access_token}',
                'Content-Type': 'application/json',
                'X-JOBBER-GRAPHQL-VERSION': API_VERSION
            },
            json={'query': test_query}
        )
        if response.status_code == 200 and 'errors' not in response.json():
            return access_token

    # Refresh token if available
    if refresh_token and client_id and client_secret:
        response = requests.post(
            'https://api.getjobber.com/api/oauth/token',
            data={
                'client_id': client_id,
                'client_secret': client_secret,
                'refresh_token': refresh_token,
                'grant_type': 'refresh_token'
            }
        )
        if response.status_code == 200:
            tokens = response.json()
            return tokens['access_token']

    raise Exception("Unable to get valid access token")


def fetch_jobs(access_token, start_date, end_date):
    """Fetch all jobs within date range from Jobber API"""
    all_jobs = []
    cursor = None

    while True:
        query = '''
        query($cursor: String) {
            jobs(first: 100, after: $cursor) {
                nodes {
                    jobNumber
                    title
                    jobStatus
                    startAt
                    endAt
                    total
                    client {
                        name
                    }
                    jobCosting {
                        totalRevenue
                        totalCost
                    }
                    invoices {
                        nodes {
                            invoiceNumber
                            total
                            invoiceStatus
                        }
                    }
                    visits {
                        nodes {
                            assignedUsers {
                                nodes {
                                    name {
                                        full
                                    }
                                }
                            }
                        }
                    }
                    customFields {
                        nodes {
                            label
                            valueText
                        }
                    }
                }
                pageInfo {
                    hasNextPage
                    endCursor
                }
            }
        }
        '''

        response = requests.post(
            API_URL,
            headers={
                'Authorization': f'Bearer {access_token}',
                'Content-Type': 'application/json',
                'X-JOBBER-GRAPHQL-VERSION': API_VERSION
            },
            json={'query': query, 'variables': {'cursor': cursor}}
        )

        if response.status_code != 200:
            print(f"API Error: {response.status_code} - {response.text}")
            break

        data = response.json()
        if 'errors' in data:
            print(f"GraphQL Error: {data['errors']}")
            break

        jobs = data['data']['jobs']['nodes']

        # Filter jobs by date and type
        for job in jobs:
            job_start = job.get('startAt')
            if job_start:
                job_date = datetime.fromisoformat(job_start.replace('Z', '+00:00'))
                if start_date <= job_date.date() <= end_date:
                    # Check if it's an enhancement job (by title or custom field)
                    title_lower = (job.get('title') or '').lower()
                    is_enhancement = 'enhancement' in title_lower
                    is_contracted = 'contracted' in title_lower and 'enhancement' in title_lower

                    # Also check custom fields for job type
                    for cf in job.get('customFields', {}).get('nodes', []):
                        if cf.get('label', '').lower() == 'job type':
                            val = (cf.get('valueText') or '').lower()
                            if 'enhancement' in val:
                                is_enhancement = True
                            if 'contracted' in val:
                                is_contracted = True

                    if is_enhancement or is_contracted:
                        job['_is_contracted'] = is_contracted
                        job['_is_enhancement'] = is_enhancement and not is_contracted
                        all_jobs.append(job)

        # Pagination
        page_info = data['data']['jobs']['pageInfo']
        if not page_info['hasNextPage']:
            break
        cursor = page_info['endCursor']

        # Rate limit protection
        import time
        time.sleep(0.5)

    return all_jobs


def fetch_invoices(access_token, job_numbers):
    """Fetch invoices for specific job numbers"""
    all_invoices = []
    cursor = None

    while True:
        query = '''
        query($cursor: String) {
            invoices(first: 100, after: $cursor) {
                nodes {
                    invoiceNumber
                    total
                    invoiceStatus
                    issuedDate
                    dueDate
                    client {
                        name
                    }
                    job {
                        jobNumber
                    }
                }
                pageInfo {
                    hasNextPage
                    endCursor
                }
            }
        }
        '''

        response = requests.post(
            API_URL,
            headers={
                'Authorization': f'Bearer {access_token}',
                'Content-Type': 'application/json',
                'X-JOBBER-GRAPHQL-VERSION': API_VERSION
            },
            json={'query': query, 'variables': {'cursor': cursor}}
        )

        if response.status_code != 200:
            break

        data = response.json()
        if 'errors' in data:
            break

        invoices = data['data']['invoices']['nodes']
        for inv in invoices:
            job = inv.get('job')
            if job and job.get('jobNumber') in job_numbers:
                all_invoices.append(inv)

        page_info = data['data']['invoices']['pageInfo']
        if not page_info['hasNextPage']:
            break
        cursor = page_info['endCursor']

        import time
        time.sleep(0.5)

    return all_invoices


def generate_report(jobs, invoices, month, year):
    """Generate Excel report from job and invoice data"""
    wb = Workbook()
    ws = wb.active
    ws.title = "Executive Dashboard"

    # Set column widths
    for col, width in {'A': 3, 'B': 18, 'C': 14, 'D': 14, 'E': 3, 'F': 18, 'G': 14, 'H': 14}.items():
        ws.column_dimensions[col].width = width

    # Title
    month_name = datetime(year, month, 1).strftime('%B')
    ws.merge_cells('B2:H2')
    ws['B2'] = f"ENHANCEMENT JOBS - {month_name.upper()} {year}"
    ws['B2'].font = title_font

    ws.merge_cells('B3:H3')
    ws['B3'] = f"Generated: {datetime.now().strftime('%B %d, %Y at %I:%M %p')}"
    ws['B3'].font = Font(size=10, italic=True, color="808080")

    # Calculate metrics
    enhancement_jobs = [j for j in jobs if j.get('_is_enhancement')]
    contracted_jobs = [j for j in jobs if j.get('_is_contracted')]

    total_revenue = sum(float(j.get('jobCosting', {}).get('totalRevenue') or j.get('total') or 0) for j in jobs)
    total_cost = sum(float(j.get('jobCosting', {}).get('totalCost') or 0) for j in jobs)
    total_profit = total_revenue - total_cost

    # KPIs
    ws['B6'] = "TOTAL JOBS"
    ws['B6'].font = Font(size=9, color="808080")
    ws['B7'] = len(jobs)
    ws['B7'].font = Font(bold=True, size=24, color=NAVY)
    ws['B8'] = f"{len(enhancement_jobs)} std / {len(contracted_jobs)} contracted"
    ws['B8'].font = Font(size=8, color="808080")

    ws['D6'] = "TOTAL REVENUE"
    ws['D6'].font = Font(size=9, color="808080")
    ws['D7'] = total_revenue
    ws['D7'].font = Font(bold=True, size=24, color=NAVY)
    ws['D7'].number_format = currency_format_whole

    ws['F6'] = "GROSS PROFIT"
    ws['F6'].font = Font(size=9, color="808080")
    ws['F7'] = total_profit
    ws['F7'].font = Font(bold=True, size=24, color=NAVY)
    ws['F7'].number_format = currency_format_whole
    margin = total_profit / total_revenue if total_revenue > 0 else 0
    ws['F8'] = f"margin: {margin:.1%}"
    ws['F8'].font = Font(size=9, bold=True, color="548235")

    # Add KPI backgrounds
    for col in [2, 3, 4, 5, 6, 7]:
        for row in range(6, 9):
            ws.cell(row=row, column=col).fill = PatternFill(start_color=LIGHT_BLUE, end_color=LIGHT_BLUE, fill_type="solid")

    # Jobs data sheet
    ws_jobs = wb.create_sheet("Jobs")
    headers = ['Job #', 'Client', 'Title', 'Status', 'Start Date', 'Revenue', 'Cost', 'Profit', 'Type']
    for col, header in enumerate(headers, 1):
        cell = ws_jobs.cell(row=1, column=col)
        cell.value = header
        cell.font = header_font
        cell.fill = header_fill
        cell.border = thin_border

    for row_idx, job in enumerate(jobs, 2):
        costing = job.get('jobCosting') or {}
        revenue = float(costing.get('totalRevenue') or job.get('total') or 0)
        cost = float(costing.get('totalCost') or 0)
        profit = revenue - cost
        job_type = 'Contracted Enhancement' if job.get('_is_contracted') else 'Enhancement'

        ws_jobs.cell(row=row_idx, column=1).value = job.get('jobNumber')
        ws_jobs.cell(row=row_idx, column=2).value = job.get('client', {}).get('name', '')
        ws_jobs.cell(row=row_idx, column=3).value = job.get('title', '')
        ws_jobs.cell(row=row_idx, column=4).value = job.get('jobStatus', '')
        ws_jobs.cell(row=row_idx, column=5).value = job.get('startAt', '')[:10] if job.get('startAt') else ''
        ws_jobs.cell(row=row_idx, column=6).value = revenue
        ws_jobs.cell(row=row_idx, column=6).number_format = currency_format
        ws_jobs.cell(row=row_idx, column=7).value = cost
        ws_jobs.cell(row=row_idx, column=7).number_format = currency_format
        ws_jobs.cell(row=row_idx, column=8).value = profit
        ws_jobs.cell(row=row_idx, column=8).number_format = currency_format
        ws_jobs.cell(row=row_idx, column=9).value = job_type

        for col in range(1, 10):
            ws_jobs.cell(row=row_idx, column=col).border = thin_border
            if row_idx % 2 == 0:
                ws_jobs.cell(row=row_idx, column=col).fill = alt_row_fill

    # Conditional formatting for profit
    if len(jobs) > 0:
        ws_jobs.conditional_formatting.add(f'H2:H{len(jobs)+1}', FormulaRule(formula=['$H2>0'], fill=green_fill))
        ws_jobs.conditional_formatting.add(f'H2:H{len(jobs)+1}', FormulaRule(formula=['$H2<0'], fill=red_fill))

    ws_jobs.freeze_panes = 'A2'

    # Auto-fit columns
    for col in ws_jobs.columns:
        max_length = max(len(str(cell.value or '')) for cell in col)
        ws_jobs.column_dimensions[col[0].column_letter].width = min(max_length + 2, 30)

    # Invoices sheet
    ws_inv = wb.create_sheet("Invoices")
    inv_headers = ['Invoice #', 'Client', 'Job #', 'Total', 'Status', 'Issued', 'Due']
    for col, header in enumerate(inv_headers, 1):
        cell = ws_inv.cell(row=1, column=col)
        cell.value = header
        cell.font = header_font
        cell.fill = header_fill
        cell.border = thin_border

    for row_idx, inv in enumerate(invoices, 2):
        ws_inv.cell(row=row_idx, column=1).value = inv.get('invoiceNumber')
        ws_inv.cell(row=row_idx, column=2).value = inv.get('client', {}).get('name', '')
        ws_inv.cell(row=row_idx, column=3).value = inv.get('job', {}).get('jobNumber', '')
        ws_inv.cell(row=row_idx, column=4).value = float(inv.get('total') or 0)
        ws_inv.cell(row=row_idx, column=4).number_format = currency_format
        ws_inv.cell(row=row_idx, column=5).value = inv.get('invoiceStatus', '')
        ws_inv.cell(row=row_idx, column=6).value = inv.get('issuedDate', '')
        ws_inv.cell(row=row_idx, column=7).value = inv.get('dueDate', '')

        for col in range(1, 8):
            ws_inv.cell(row=row_idx, column=col).border = thin_border

    ws_inv.freeze_panes = 'A2'

    return wb


def main():
    print("Enhancement Report Generator")
    print("=" * 50)

    # Check for month override from environment or command line
    import sys
    month_override = os.environ.get('REPORT_MONTH') or (sys.argv[1] if len(sys.argv) > 1 else None)
    use_csv = os.environ.get('USE_CSV', '').lower() == 'true' or '--csv' in sys.argv

    # Get current date info
    today = datetime.now()

    # Determine report month
    if month_override:
        report_month = int(month_override)
        report_year = today.year
        # If requesting a future month, assume previous year
        if report_month > today.month:
            report_year -= 1
    else:
        report_month = today.month
        report_year = today.year

    # Date range for target month
    start_date = datetime(report_year, report_month, 1).date()
    if report_month == 12:
        end_date = datetime(report_year + 1, 1, 1).date() - timedelta(days=1)
    else:
        end_date = datetime(report_year, report_month + 1, 1).date() - timedelta(days=1)

    print(f"Report Period: {start_date} to {end_date}")

    jobs = []
    invoices = []

    # Try API first (if not forcing CSV mode)
    if not use_csv:
        try:
            print("Attempting Jobber API...")
            access_token = get_access_token()
            print("Authentication successful!")

            print("Fetching enhancement jobs from API...")
            jobs = fetch_jobs(access_token, start_date, end_date)
            print(f"Found {len(jobs)} enhancement jobs from API")

            if jobs:
                job_numbers = {j.get('jobNumber') for j in jobs if j.get('jobNumber')}
                print("Fetching related invoices from API...")
                invoices = fetch_invoices(access_token, job_numbers)
                print(f"Found {len(invoices)} related invoices")

        except Exception as e:
            print(f"API access failed: {e}")
            print("Falling back to CSV mode...")
            jobs = []

    # Use CSV fallback if API returned no data or CSV mode is forced
    if not jobs:
        print("\nUsing CSV fallback mode...")

        # Find latest CSV files
        jobs_csv = find_latest_csv(JOBS_CSV_PATTERN)
        invoices_csv = find_latest_csv(INVOICES_CSV_PATTERN)

        if not jobs_csv:
            print(f"ERROR: No jobs CSV found matching pattern: {JOBS_CSV_PATTERN}")
            print(f"Please export 'One-off jobs' report from Jobber to: {DOWNLOADS_DIR}")
            return None

        print(f"Loading jobs from: {os.path.basename(jobs_csv)}")
        jobs = load_jobs_from_csv(jobs_csv, start_date, end_date)
        print(f"Found {len(jobs)} enhancement jobs in date range")

        if jobs and invoices_csv:
            job_numbers = {j.get('jobNumber') for j in jobs if j.get('jobNumber')}
            print(f"Loading invoices from: {os.path.basename(invoices_csv)}")
            invoices = load_invoices_from_csv(invoices_csv, job_numbers)
            print(f"Found {len(invoices)} related invoices")

    if not jobs:
        print("\nNo enhancement jobs found for this period.")
        print("Please ensure:")
        print("  1. Jobs exist in Jobber with Job Type = 'Enhancement' or 'Contracted Enhancement'")
        print("  2. Jobs have Scheduled start dates within the report period")
        print("  3. CSV exports are up-to-date in Downloads folder")
        return None

    # Generate report
    print("\nGenerating Excel report...")
    wb = generate_report(jobs, invoices, report_month, report_year)

    # Save report (xlsx format - openpyxl doesn't support xlsm)
    month_name = datetime(report_year, report_month, 1).strftime('%B')
    filename = f"{report_month}-OneOffReport-{month_name}.xlsx"
    wb.save(filename)
    print(f"Report saved: {filename}")

    # Print summary
    enhancement_count = sum(1 for j in jobs if j.get('_is_enhancement'))
    contracted_count = sum(1 for j in jobs if j.get('_is_contracted'))
    total_revenue = sum(float(j.get('jobCosting', {}).get('totalRevenue') or j.get('total') or 0) for j in jobs)

    print(f"\n{'='*50}")
    print(f"REPORT SUMMARY - {month_name} {report_year}")
    print(f"{'='*50}")
    print(f"Enhancement Jobs: {enhancement_count}")
    print(f"Contracted Enhancement Jobs: {contracted_count}")
    print(f"Total Jobs: {len(jobs)}")
    print(f"Total Revenue: ${total_revenue:,.2f}")
    print(f"Related Invoices: {len(invoices)}")

    return filename


if __name__ == "__main__":
    main()
