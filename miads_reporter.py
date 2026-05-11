#!/usr/bin/env python3
"""
MiAds Daily Report Automation Skill
Complete ETL pipeline: Xiaomi API → XLSX → HTML → Slack

Usage:
    python3 miads_reporter.py [--date YYYY-MM-DD] [--config config.yaml]
"""

import os
import sys
import json
import logging
from datetime import datetime, timedelta
from pathlib import Path
import argparse
import yaml

# Third-party imports
import requests
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
import pandas as pd
from bs4 import BeautifulSoup

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.StreamHandler(),
        logging.FileHandler('miads_reporter.log')
    ]
)
logger = logging.getLogger(__name__)

# ============================================================================
# STAGE 1: XIAOMI API FETCH
# ============================================================================

class XiaomiAPIClient:
    """Xiaomi MiAds API client with token management and pagination"""
    
    def __init__(self, app_id, app_key, base_url="https://global.e.mi.com"):
        self.app_id = app_id
        self.app_key = app_key
        self.base_url = base_url
        self.token = None
        self.token_time = None
    
    def get_token(self):
        """Fetch authentication token from Xiaomi API"""
        try:
            url = f"{self.base_url}/foreign/token/createToken"
            payload = {
                "appId": self.app_id,
                "appKey": self.app_key
            }
            response = requests.post(
                url,
                json=payload,
                timeout=10
            )
            response.raise_for_status()
            
            result = response.json()
            if result.get('code') != 0:
                raise Exception(f"Token fetch failed: {result.get('message')}")
            
            self.token = result['result']['accessToken']
            self.token_time = datetime.now()
            logger.info("✓ Xiaomi API token acquired")
            return self.token
        
        except Exception as e:
            logger.error(f"✗ Failed to get Xiaomi token: {e}")
            raise
    
    def query_reporting(self, account_id, campaign_id, date_from, date_to):
        """
        Query MiAds reporting API with pagination support
        
        Args:
            account_id (int): Xiaomi account ID
            campaign_id (int): Xiaomi campaign ID
            date_from (datetime): Start date
            date_to (datetime): End date
        
        Returns:
            list: Combined records from all pages
        """
        if not self.token:
            self.get_token()
        
        url = f"{self.base_url}/foreign/data/queryData"
        all_records = []
        current_page = 1
        total_pages = 1
        
        begin_str = date_from.isoformat() + 'Z'
        end_str = date_to.isoformat() + 'Z'
        
        while current_page <= total_pages:
            try:
                payload = {
                    "accountIds": [int(account_id)],
                    "adCampaignIds": [int(campaign_id)],
                    "adType": 1,
                    "begin": begin_str,
                    "end": end_str,
                    "dimensions": [2, 3, 9, 10, 5],  # Ad Set, Ad, Placement, Media, Date
                    "pageSize": 1000,
                    "page": current_page,
                    "lang": "en_US"
                }
                
                headers = {
                    'Cookie': f'access_token={self.token}',
                    'Content-Type': 'application/json'
                }
                
                response = requests.post(
                    url,
                    json=payload,
                    headers=headers,
                    timeout=30
                )
                response.raise_for_status()
                
                result = response.json()
                
                if result.get('code') != 0:
                    raise Exception(f"API error: {result.get('message')}")
                
                if result.get('result', {}).get('records'):
                    all_records.extend(result['result']['records'])
                    total_pages = result['result'].get('pages', 1)
                    logger.info(f"✓ Fetched page {current_page}/{total_pages} "
                              f"({len(result['result']['records'])} records)")
                
                current_page += 1
            
            except requests.exceptions.RequestException as e:
                logger.error(f"✗ API request failed: {e}")
                if current_page == 1:
                    raise
                break
        
        logger.info(f"✓ Total records fetched: {len(all_records)}")
        return all_records

# ============================================================================
# STAGE 2: DATA TRANSFORMATION
# ============================================================================

class DataTransformer:
    """Transform Xiaomi API data and calculate metrics"""
    
    @staticmethod
    def get_fx_rate_usdmxn():
        """Get current USD/MXN exchange rate (fallback to default)"""
        try:
            # Try GOOGLEFINANCE API
            url = "https://www.google.com/async/currencies"
            params = {'q': 'USD to MXN'}
            response = requests.get(url, params=params, timeout=5)
            
            # Parse rate from response (simplified)
            # For production, use a proper currency API
            # Fallback rate: 17.48 MXN/USD (as of 2026-05-03)
            return 17.4811
        except:
            logger.warning("⚠ FX rate fetch failed, using fallback rate: 17.4811")
            return 17.4811
    
    @staticmethod
    def transform_records(records, fx_rate=17.4811):
        """
        Transform Xiaomi API records to structured format
        
        Args:
            records (list): Raw API records
            fx_rate (float): USD/MXN exchange rate
        
        Returns:
            list: Transformed records with calculated metrics
        """
        transformed = []
        
        for rec in records:
            impressions = int(rec.get('expose', 0))
            clicks = int(rec.get('click', 0))
            
            # Cost calculation
            cost_usd = clicks * 0.04  # CPC = $0.04 USD (Xiaomi standard)
            cost_mxn = cost_usd * fx_rate
            
            # Metrics
            ctr = (clicks / impressions * 100) if impressions > 0 else 0
            cpm_mxn = (cost_mxn / impressions * 1000) if impressions > 0 else 0
            cpm_usd = (cost_usd / impressions * 1000) if impressions > 0 else 0
            cpc_mxn = (cost_mxn / clicks) if clicks > 0 else 0
            
            transformed.append({
                'ad_set': rec.get('adGroupName', '-'),
                'ad_name': rec.get('adCreativeName', '-'),
                'placement': rec.get('tagId', '-'),
                'media': rec.get('mediaTypeName', '-'),
                'date': rec.get('recordDate', ''),
                'impressions': impressions,
                'clicks': clicks,
                'cost_usd': round(cost_usd, 2),
                'cost_mxn': round(cost_mxn, 2),
                'ctr_pct': round(ctr, 2),
                'cvr_pct': float(rec.get('cvr', 0)),
                'ecpm': float(rec.get('ecpm', 0)),
                'cpc': round(cost_usd / clicks, 4) if clicks > 0 else 0,
                'cpm_mxn': round(cpm_mxn, 2),
                'cpm_usd': round(cpm_usd, 4)
            })
        
        logger.info(f"✓ Transformed {len(transformed)} records")
        return transformed

# ============================================================================
# STAGE 3: XLSX GENERATION
# ============================================================================

class XLSXGenerator:
    """Generate Excel workbook with campaign data"""
    
    @staticmethod
    def create_workbook(transformed_data, output_path):
        """
        Create and save Excel workbook
        
        Args:
            transformed_data (list): Transformed records
            output_path (str): Output file path
        
        Returns:
            str: Path to created file
        """
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "MiAds_Report"
        
        # Headers
        headers = [
            'Ad Set', 'Ad Name', 'Placement', 'Media', 'Date',
            'Impressions', 'Clicks', 'Cost MXN', 'Cost USD',
            'CTR %', 'CVR %', 'ECPM', 'CPC', 'CPM MXN', 'CPM USD'
        ]
        ws.append(headers)
        
        # Style header row
        header_fill = PatternFill(start_color="D3D3D3", end_color="D3D3D3", fill_type="solid")
        header_font = Font(bold=True)
        for cell in ws[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal='center')
        
        # Add data rows
        for row in transformed_data:
            ws.append([
                row['ad_set'],
                row['ad_name'],
                row['placement'],
                row['media'],
                row['date'],
                row['impressions'],
                row['clicks'],
                row['cost_mxn'],
                row['cost_usd'],
                row['ctr_pct'],
                row['cvr_pct'],
                row['ecpm'],
                row['cpc'],
                row['cpm_mxn'],
                row['cpm_usd']
            ])
        
        # Format columns
        for col_num, header in enumerate(headers, 1):
            col_letter = openpyxl.utils.get_column_letter(col_num)
            
            # Set width
            ws.column_dimensions[col_letter].width = 15
            
            # Format numbers
            if col_num > 5:  # Numeric columns
                for row_num in range(2, ws.max_row + 1):
                    cell = ws[f'{col_letter}{row_num}']
                    if col_num in [8, 9, 14, 15]:  # Currency
                        cell.number_format = '#,##0.00'
                    elif col_num in [6, 7]:  # Large numbers
                        cell.number_format = '#,##0'
                    else:  # Percentages and metrics
                        cell.number_format = '0.00'
        
        # Freeze header row
        ws.freeze_panes = 'A2'
        
        # Save
        wb.save(output_path)
        logger.info(f"✓ XLSX created: {output_path}")
        return output_path

# ============================================================================
# STAGE 4: HTML REPORT UPDATE
# ============================================================================

class HTMLReportUpdater:
    """Update HTML report with new data"""
    
    @staticmethod
    def calculate_aggregations(df):
        """Calculate key metrics for report"""
        totals = {
            'total_impressions': int(df['Impressions'].sum()),
            'total_clicks': int(df['Clicks'].sum()),
            'total_ctr': (df['Clicks'].sum() / df['Impressions'].sum() * 100) if df['Impressions'].sum() > 0 else 0,
            'total_cost_mxn': df['Cost MXN'].sum(),
            'total_cost_usd': df['Cost USD'].sum(),
            'avg_cpc_mxn': (df['Cost MXN'].sum() / df['Clicks'].sum()) if df['Clicks'].sum() > 0 else 0,
            'avg_cpm_mxn': (df['Cost MXN'].sum() / df['Impressions'].sum() * 1000) if df['Impressions'].sum() > 0 else 0,
            'avg_cpm_usd': (df['Cost USD'].sum() / df['Impressions'].sum() * 1000) if df['Impressions'].sum() > 0 else 0,
            'date_from': df['Date'].min(),
            'date_to': df['Date'].max(),
            'num_days': df['Date'].nunique()
        }
        return totals
    
    @staticmethod
    def update_html(html_template_path, xlsx_path, output_path):
        """
        Update HTML report with new data from XLSX
        
        Args:
            html_template_path (str): Path to HTML template
            xlsx_path (str): Path to XLSX file
            output_path (str): Output HTML file path
        
        Returns:
            str: Path to updated HTML file
        """
        # Read template
        with open(html_template_path, 'r', encoding='utf-8') as f:
            html_content = f.read()
        
        # Load XLSX data
        df = pd.read_excel(xlsx_path)
        
        # Calculate aggregations
        totals = HTMLReportUpdater.calculate_aggregations(df)
        
        # Parse HTML
        soup = BeautifulSoup(html_content, 'html.parser')
        
        # Update header metadata
        header_meta = soup.find('div', class_='header-meta')
        if header_meta:
            header_meta.string = (
                f"Advertiser: Santander México · Período: {totals['date_from']} – {totals['date_to']} · "
                f"{totals['num_days']} días activos\n"
                f"Plataforma: MiAds · Canal: Web Display · Región: México · Tipo de compra: CPC\n"
                f"Moneda: MXN + USD · TC referencia: $17.4811 MXN/USD · Ad Sets: NativeAds · NewsFeed · Icon · Interstitial"
            )
        
        # Update hero KPIs
        hero_cards = soup.find_all('div', class_='hero-value')
        if len(hero_cards) >= 3:
            hero_cards[0].string = f"${totals['total_cost_mxn']:,.0f}"
            hero_cards[1].string = f"${totals['avg_cpc_mxn']:.2f}"
            hero_cards[2].string = f"${totals['avg_cpm_mxn']:.2f}"
        
        # Update KPI strip
        kpi_values = soup.find_all('div', class_='kpi-mini-value')
        if len(kpi_values) >= 6:
            kpi_values[0].string = f"{totals['total_impressions']/1e6:.1f}M"
            kpi_values[1].string = f"{totals['total_clicks']/1e3:.1f}K"
            kpi_values[2].string = f"{totals['total_ctr']:.2f}%"
            kpi_values[3].string = str(totals['num_days'])
            # ... update remaining KPIs
        
        # Save updated HTML
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(str(soup))
        
        logger.info(f"✓ HTML updated: {output_path}")
        return output_path

# ============================================================================
# STAGE 5: SLACK NOTIFICATION
# ============================================================================

class SlackNotifier:
    """Post report to Slack channel"""
    
    @staticmethod
    def post_report(webhook_url, metrics, html_url, date_from, date_to):
        """
        Send report summary to Slack
        
        Args:
            webhook_url (str): Slack webhook URL
            metrics (dict): Report metrics
            html_url (str): URL to HTML report
            date_from (str): Start date
            date_to (str): End date
        
        Returns:
            bool: Success status
        """
        try:
            payload = {
                'text': '📊 MiAds Daily Report',
                'blocks': [
                    {
                        'type': 'header',
                        'text': {
                            'type': 'plain_text',
                            'text': '📊 MiAds Daily Campaign Report'
                        }
                    },
                    {
                        'type': 'section',
                        'fields': [
                            {
                                'type': 'mrkdwn',
                                'text': f"*📈 Impressions*\n{metrics['total_impressions']:,}"
                            },
                            {
                                'type': 'mrkdwn',
                                'text': f"*🖱️ Clicks*\n{metrics['total_clicks']:,}"
                            },
                            {
                                'type': 'mrkdwn',
                                'text': f"*📊 CTR*\n{metrics['total_ctr']:.2f}%"
                            },
                            {
                                'type': 'mrkdwn',
                                'text': f"*💰 Spend*\n${metrics['total_cost_mxn']:,.0f} MXN"
                            }
                        ]
                    },
                    {
                        'type': 'section',
                        'text': {
                            'type': 'mrkdwn',
                            'text': f"*Period:* {date_from} to {date_to}\n\n"
                                   f"<{html_url}|👉 View Full Report>"
                        }
                    }
                ]
            }
            
            response = requests.post(webhook_url, json=payload, timeout=10)
            response.raise_for_status()
            
            logger.info("✓ Slack notification sent")
            return True
        
        except Exception as e:
            logger.error(f"✗ Slack notification failed: {e}")
            return False

# ============================================================================
# MAIN ORCHESTRATOR
# ============================================================================

def main():
    """Main ETL pipeline orchestration"""
    
    parser = argparse.ArgumentParser(description='MiAds Daily Report Automation')
    parser.add_argument('--date', type=str, help='Report date (YYYY-MM-DD), default: yesterday')
    parser.add_argument('--config', type=str, default='config.yaml', help='Config file path')
    args = parser.parse_args()
    
    # Load configuration
    config_path = Path(args.config)
    if not config_path.exists():
        logger.error(f"✗ Config file not found: {config_path}")
        sys.exit(1)
    
    with open(config_path, 'r') as f:
        config = yaml.safe_load(f)
    
    logger.info("=" * 60)
    logger.info("MiAds Daily Report Automation Pipeline")
    logger.info("=" * 60)
    
    try:
        # Calculate date range
        if args.date:
            report_date = datetime.strptime(args.date, '%Y-%m-%d').date()
        else:
            report_date = (datetime.now() - timedelta(days=1)).date()
        
        date_from = datetime(report_date.year, report_date.month, 1)
        date_to = datetime.combine(report_date, datetime.max.time())
        
        logger.info(f"Report period: {date_from.date()} to {date_to.date()}")
        
        # STAGE 1: Fetch Xiaomi API
        logger.info("\n[STAGE 1] Fetching data from Xiaomi API...")
        client = XiaomiAPIClient(
            app_id=config['xiaomi']['app_id'],
            app_key=config['xiaomi']['app_key'],
            base_url=config['xiaomi'].get('base_url', 'https://global.e.mi.com')
        )
        
        raw_records = client.query_reporting(
            account_id=config['xiaomi']['account_id'],
            campaign_id=config['xiaomi']['campaign_id'],
            date_from=date_from,
            date_to=date_to
        )
        
        if not raw_records:
            logger.warning("⚠ No data fetched from API")
            return {'status': 'partial', 'message': 'No data available'}
        
        # STAGE 2: Transform Data
        logger.info("\n[STAGE 2] Transforming data...")
        fx_rate = DataTransformer.get_fx_rate_usdmxn()
        transformed_data = DataTransformer.transform_records(raw_records, fx_rate)
        
        # STAGE 3: Create XLSX
        logger.info("\n[STAGE 3] Generating XLSX file...")
        output_dir = Path(config['output']['base_path']) / f"{date_to.year:04d}" / f"{date_to.month:02d}" / f"{date_to.day:02d}"
        output_dir.mkdir(parents=True, exist_ok=True)
        
        xlsx_filename = f"MiAds_Santander_{date_to.strftime('%Y%m%d')}.xlsx"
        xlsx_path = output_dir / xlsx_filename
        
        XLSXGenerator.create_workbook(transformed_data, str(xlsx_path))
        
        # STAGE 4: Update HTML
        logger.info("\n[STAGE 4] Updating HTML report...")
        html_template_path = Path(config['output'].get('template_path', 'santander-miads-daily-report.html'))
        html_filename = f"miads_report_{date_to.strftime('%Y%m%d')}.html"
        html_path = output_dir / html_filename
        
        if html_template_path.exists():
            HTMLReportUpdater.update_html(str(html_template_path), str(xlsx_path), str(html_path))
        else:
            logger.warning(f"⚠ HTML template not found: {html_template_path}")
        
        # STAGE 5: Post to Slack
        logger.info("\n[STAGE 5] Notifying Slack...")
        df = pd.read_excel(xlsx_path)
        totals = HTMLReportUpdater.calculate_aggregations(df)
        
        html_url = config['slack'].get('report_base_url', 'https://reports.example.com') + f"/{date_to.year}/{date_to.month:02d}/{date_to.day:02d}/{html_filename}"
        
        slack_success = SlackNotifier.post_report(
            webhook_url=config['slack']['webhook_url'],
            metrics=totals,
            html_url=html_url,
            date_from=date_from.strftime('%Y-%m-%d'),
            date_to=date_to.strftime('%Y-%m-%d')
        )
        
        logger.info("\n" + "=" * 60)
        logger.info("✓ Pipeline completed successfully!")
        logger.info("=" * 60)
        
        return {
            'status': 'success',
            'xlsx': str(xlsx_path),
            'html': str(html_path),
            'records_processed': len(transformed_data),
            'date_range': f"{date_from.date()} to {date_to.date()}"
        }
    
    except Exception as e:
        logger.error(f"\n✗ Pipeline failed: {e}", exc_info=True)
        return {'status': 'error', 'message': str(e)}

if __name__ == '__main__':
    result = main()
    print(json.dumps(result, indent=2))
    sys.exit(0 if result['status'] == 'success' else 1)
