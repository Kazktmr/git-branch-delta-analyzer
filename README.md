# Git Branch Delta Analyzer

**A powerful TPM tool to analyze and visualize deltas between two Git branches or time snapshots.**

This tool helps Technical Program Managers (TPMs), engineering leads, and developers understand exactly what changed between two similar branches, commits, or time periods — with rich Excel reports and beautiful interactive HTML dashboards.

## ✨ Key Features

- Compare any two Git references: branches, commit hashes, time snapshots (`main@{7.days.ago}`), etc.
- Smart JIRA ticket extraction from commit messages
- Detailed Excel report (Commits + Files Changed + Metadata)
- Beautiful interactive HTML dashboard with multiple tabs and charts
- **Highlighted Paths** feature — focus deep analysis on critical files/folders
- Always-available latest example report in `example_reports/latest/`
- GitHub Actions automation for fresh demo reports

## 📥 Installation

1. Clone this repository:
   ```bash
   git clone https://github.com/Kazktmr/git-branch-delta-analyzer.git
   cd git-branch-delta-analyzer
   ```

2. Install dependencies:
   ```bash
   pip install gitpython pandas openpyxl plotly
   ```

## 🚀 Quick Start

### Time-based comparison (most common for TPMs)
```bash
python branch_delta_analyzer.py --github-repo openclaw/openclaw --base main --days-ago 7
```

### Compare two branches
```bash
python branch_delta_analyzer.py --repo-path /path/to/your/repo --base main --compare feature/new-ui
```

### Using commit hashes
```bash
python branch_delta_analyzer.py --github-repo username/repo --base abc1234 --compare def5678
```

## Full Documentation

See the detailed sections below for all options, configuration, and usage.

## Configuration

### Highlighted Paths (`highlight_config.json`)
Create a `highlight_config.json` file to get deep insights on specific paths:

```json
{
  "highlighted_paths": [
    "src/core/",
    "src/plugins/",
    "package.json"
  ],
  "include_subpaths": true
}
```

## Reports Generated

- **Excel**: `delta_summary_YYYYMMDD.xlsx` with multiple sheets (Commits, Files_Changed, Metadata)
- **HTML**: Interactive dashboard with KPI cards, charts, Metadata tab, Highlighted Paths tab, etc.

## Example Reports

The latest example report is always available here:
- [LATEST Interactive Report](example_reports/latest/LATEST_delta_report.html)
- [LATEST Excel Report](example_reports/latest/LATEST_delta_summary.xlsx)

## Automation

GitHub Actions automatically generates fresh example reports on push and daily.

---

**Made for TPMs | Clean | Visual | Actionable**
