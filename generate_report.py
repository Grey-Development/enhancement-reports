"""
Enhancement Jobs Report Generator
Pulls data from Jobber API and generates monthly Excel report
Runs weekly on Fridays via GitHub Actions
"""

import os
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

    # Get access token
    print("Authenticating with Jobber API...")
    access_token = get_access_token()
    print("Authentication successful!")

    # Fetch jobs
    print("Fetching enhancement jobs...")
    jobs = fetch_jobs(access_token, start_date, end_date)
    print(f"Found {len(jobs)} enhancement jobs")

    # Fetch invoices
    job_numbers = {j.get('jobNumber') for j in jobs if j.get('jobNumber')}
    print("Fetching related invoices...")
    invoices = fetch_invoices(access_token, job_numbers)
    print(f"Found {len(invoices)} related invoices")

    # Generate report
    print("Generating Excel report...")
    wb = generate_report(jobs, invoices, report_month, report_year)

    # Save report (xlsx format - openpyxl doesn't support xlsm)
    month_name = datetime(report_year, report_month, 1).strftime('%B')
    filename = f"{report_month}-OneOffReport-{month_name}.xlsx"
    wb.save(filename)
    print(f"Report saved: {filename}")

    return filename


if __name__ == "__main__":
    main()
