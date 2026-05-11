# MiAds Daily Report Automation Skill

**Automate daily MiAds campaign reporting with API integration, XLSX generation, and Slack distribution.**

## 🚀 Quick Start (5 minutes)

### 1. Install Dependencies

```bash
pip install requests openpyxl pandas beautifulsoup4 pyyaml
```

### 2. Configure

Create `config.yaml`:

```yaml
xiaomi:
  app_id: "XILANY1_S_A"
  app_key: "957a96f9-a058-45da-90cf-a7dd1b491466"
  account_id: 18348
  campaign_id: 276730

slack:
  webhook_url: "https://hooks.slack.com/services/YOUR/WEBHOOK/URL"
  
output:
  base_path: "/outputs/miads_reports"
```

### 3. Run

```bash
python3 miads_reporter.py
```

**Output:**
```
✓ XLSX created: /outputs/miads_reports/2026/05/03/MiAds_Santander_20260503.xlsx
✓ HTML updated: /outputs/miads_reports/2026/05/03/miads_report_20260503.html
✓ Slack notification sent
```

---

## 📋 What It Does

### Pipeline Stages

```
[Xiaomi API] → [Transform Data] → [Create XLSX] → [Update HTML] → [Slack]
   1. Fetch       2. Calculate      3. Generate    4. Inject       5. Notify
   Daily data     USD/MXN costs     Excel file     into template   Channel
```

### Automatic Calculations

- **Cost USD** = Clicks × $0.04 (CPC)
- **Cost MXN** = Cost USD × FX Rate
- **CTR** = (Clicks / Impressions) × 100
- **CPM** = (Cost / Impressions) × 1000
- **CPC** = Cost / Clicks

### Files Generated

```
/outputs/miads_reports/
  2026/
    05/
      03/
        ├── MiAds_Santander_20260503.xlsx  ← Data file
        └── miads_report_20260503.html     ← Report (5 sections)
```

---

## 📊 Report Sections

1. **Performance Diario** — Daily trend cards + tables + charts
2. **Ad Sets** — Compare cards + performance breakdown
3. **Top Media** — Top 10 placements ranked by clicks & CTR
4. **Distribución de Inversión** — Spend by ad set and day
5. **Recomendaciones** — Escalar / Monitorear / Pausar insights

---

## 🔄 Scheduling

### Linux Cron (Simplest)

```bash
# Daily at 8 AM UTC
0 8 * * * cd /path/to/skill && python3 miads_reporter.py
```

### Systemd Timer (Recommended)

```bash
sudo systemctl enable miads-daily-report.timer
sudo systemctl start miads-daily-report.timer
```

### Docker (Cloud)

```bash
docker build -t miads-daily-report .
docker run -e XIAOMI_APP_KEY="..." -v /outputs:/outputs miads-daily-report
```

### Google Cloud Scheduler / AWS Lambda

See `DEPLOYMENT.md` for detailed setup.

---

## 📱 Slack Integration

**Daily message includes:**
- Impressions, Clicks, CTR, Spend (MXN + USD)
- Report link
- Period covered
- Button to view full HTML

**Example:**
```
📊 MiAds Daily Report

📈 Impressions: 3.35M
🖱️ Clicks: 30,526 (0.91% CTR)
💰 Spend: $21,319 MXN / $1,221 USD
⏰ Period: 2026-05-01 to 2026-05-03

👉 [View Full Report]
```

---

## 🔧 Configuration

### Minimal Config

```yaml
xiaomi:
  app_key: "YOUR_KEY"
  account_id: 18348
  campaign_id: 276730

slack:
  webhook_url: "YOUR_WEBHOOK"
```

### Full Config (See `config.yaml.template`)

```yaml
xiaomi:
  app_id: "XILANY1_S_A"
  app_key: "957a96f9-a058-45da-90cf-a7dd1b491466"
  base_url: "https://global.e.mi.com"
  account_id: 18348
  campaign_id: 276730

slack:
  webhook_url: "https://hooks.slack.com/services/..."
  channel_id: "C0XXXXXXX"
  report_base_url: "https://reports.yourcompany.com"

output:
  base_path: "/outputs/miads_reports"
  template_path: "/path/to/template.html"
  timezone: "UTC"
```

---

## 🧪 Testing

### Test API Connection

```bash
python3 -c "
from miads_reporter import XiaomiAPIClient
client = XiaomiAPIClient('XILANY1_S_A', 'YOUR_APP_KEY')
token = client.get_token()
print('✓ Token acquired')
"
```

### Test Full Pipeline

```bash
# Run for specific date
python3 miads_reporter.py --date 2026-05-03

# Or yesterday (default)
python3 miads_reporter.py
```

### Check Output

```bash
# Verify XLSX created
ls -lh /outputs/miads_reports/2026/05/03/*.xlsx

# View HTML in browser
open /outputs/miads_reports/2026/05/03/miads_report_20260503.html

# Check logs
tail -f miads_reporter.log
```

---

## 📚 Files in Skill

```
miads-daily-report/
├── miads_reporter.py       ← Main ETL pipeline (1000+ lines)
├── SKILL.md                ← Technical documentation
├── DEPLOYMENT.md           ← Scheduling & setup guide
├── README.md               ← This file
├── config.yaml.template    ← Config template
├── requirements.txt        ← Python dependencies
└── Dockerfile              ← Container image
```

---

## 🚨 Troubleshooting

| Issue | Fix |
|-------|-----|
| "Invalid token" | Check `app_id`, `app_key` in config |
| "No data fetched" | Verify `account_id`, `campaign_id` are correct |
| "Slack failed" | Test webhook URL; check channel exists |
| "FX rate timeout" | Fallback rate used; check network |
| Cron not running | `crontab -l` to verify; check logs |

See `DEPLOYMENT.md` for detailed troubleshooting.

---

## 🔐 Security

1. **Never commit credentials** — Use `.env` or environment variables
2. **Rotate app keys** — Update every 90 days
3. **Restrict file permissions** — `chmod 600 config.yaml`
4. **Monitor logs** — Ensure no tokens/keys are logged
5. **Use VPN** — For sensitive API calls if possible

---

## 📈 Performance

- **Typical runtime**: 15-20 seconds
- **Memory usage**: ~100 MB
- **Network calls**: 2-3 requests
- **Disk I/O**: 1-2 MB per report

---

## 🔮 Future Enhancements

- [ ] Multi-account support (batch reports)
- [ ] Email distribution fallback
- [ ] Database metrics logging
- [ ] ML forecasting in recommendations
- [ ] Automated optimization (exclude underperformers)
- [ ] Weekly/monthly rollups

---

## 📞 Support

**Documentation:**
- `SKILL.md` — Architecture & implementation
- `DEPLOYMENT.md` — Scheduling & troubleshooting
- `config.yaml.template` — Configuration reference

**Logs:**
```bash
tail -f miads_reporter.log
```

**Test:**
```bash
python3 miads_reporter.py --date $(date -d yesterday +%Y-%m-%d)
```

---

## 📄 License

MIT License — See LICENSE.txt

---

## 🎯 Next Steps

1. **Install**: `pip install -r requirements.txt`
2. **Configure**: Edit `config.yaml` with your credentials
3. **Test**: `python3 miads_reporter.py --date 2026-05-03`
4. **Schedule**: Add cron job or deploy to cloud
5. **Monitor**: Check logs and Slack messages daily

**Ready to automate!** 🚀
