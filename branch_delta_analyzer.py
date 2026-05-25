#!/usr/bin/env python3
"""
Git Branch Delta Analyzer for TPMs
Compares two branches (local or remote GitHub) and generates Excel + HTML reports.
"""

import argparse
import os
import re
import sys
import tempfile
import shutil
from datetime import datetime
from pathlib import Path

try:
    import git
    import pandas as pd
    from plotly.subplots import make_subplots
    import plotly.graph_objects as go
except ImportError as e:
    print(f"❌ Missing dependency: {e}")
    print("Run: pip install gitpython pandas openpyxl plotly")
    sys.exit(1)


def parse_jira_tickets(message: str) -> list:
    """Extract JIRA-style tickets like PROJ-123"""
    return re.findall(r'([A-Z][A-Z0-9]*-\d+)', message.upper())


def clone_repo_if_needed(github_repo: str = None, repo_path: str = None) -> str:
    """Handle remote GitHub repo by cloning to temp dir"""
    if github_repo:
        print(f"📥 Cloning {github_repo}...")
        temp_dir = tempfile.mkdtemp(prefix="git_delta_")
        try:
            git.Repo.clone_from(f"https://github.com/{github_repo}.git", temp_dir)
            return temp_dir
        except Exception as e:
            print(f"❌ Failed to clone {github_repo}: {e}")
            if os.path.exists(temp_dir):
                shutil.rmtree(temp_dir, ignore_errors=True)
            sys.exit(1)
    return repo_path


def analyze_branches(repo_path: str, base_branch: str, compare_branch: str, output_dir: str = "delta_reports"):
    """Main analysis function with robust error handling"""
    try:
        repo = git.Repo(repo_path)
        
        # Validate branches exist
        try:
            repo.git.rev_parse(base_branch)
            repo.git.rev_parse(compare_branch)
        except git.exc.GitCommandError as e:
            print(f"❌ Branch error: {e}")
            print("Available branches:")
            print(repo.git.branch("-a"))
            return None, None

        print(f"🔍 Analyzing delta: {base_branch} .. {compare_branch}")
        
        # Get commits unique to compare branch
        commits = list(repo.iter_commits(f'{base_branch}..{compare_branch}'))
        
        if not commits:
            print("⚠️ No differences found between branches.")
            return None, None

        data = []
        file_changes = {}

        for commit in commits:
            jiras = parse_jira_tickets(commit.message)
            jira_str = ', '.join(jiras) if jiras else ''
            
            # File stats
            try:
                for file, stats in commit.stats.files.items():
                    if file not in file_changes:
                        file_changes[file] = {'commits': 0, 'additions': 0, 'deletions': 0}
                    file_changes[file]['commits'] += 1
                    file_changes[file]['additions'] += stats.get('insertions', 0)
                    file_changes[file]['deletions'] += stats.get('deletions', 0)
            except Exception as stats_err:
                print(f"⚠️ Could not get stats for commit {commit.hexsha[:8]}: {stats_err}")

            data.append({
                'Commit_SHA': commit.hexsha[:8],
                'Author': str(commit.author.name),
                'Date': commit.committed_datetime.strftime('%Y-%m-%d %H:%M'),
                'Message': commit.message.strip().replace('\n', ' ')[:150],
                'JIRA_Tickets': jira_str,
                'Files_Changed': len(commit.stats.files),
                'Lines_Added': sum(s.get('insertions', 0) for s in commit.stats.files.values()),
                'Lines_Deleted': sum(s.get('deletions', 0) for s in commit.stats.files.values()),
            })

        df_commits = pd.DataFrame(data)
        
        # File-level summary
        file_data = []
        for f, stats in file_changes.items():
            file_data.append({
                'File_Path': f,
                'Commits_Touching': stats['commits'],
                'Lines_Added': stats['additions'],
                'Lines_Deleted': stats['deletions'],
                'Net_Change': stats['additions'] - stats['deletions']
            })
        df_files = pd.DataFrame(file_data).sort_values('Commits_Touching', ascending=False)

        # Save outputs
        os.makedirs(output_dir, exist_ok=True)
        timestamp = datetime.now().strftime('%Y%m%d_%H%M')
        
        excel_path = os.path.join(output_dir, f'delta_summary_{timestamp}.xlsx')
        with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
            df_commits.to_excel(writer, sheet_name='Commits', index=False)
            df_files.to_excel(writer, sheet_name='Files_Changed', index=False)
        
        # Create simple HTML report
        html_path = os.path.join(output_dir, f'delta_report_{timestamp}.html')
        create_html_report(df_commits, df_files, html_path, base_branch, compare_branch)
        
        print(f"✅ Excel saved → {excel_path}")
        print(f"✅ HTML report saved → {html_path}")
        print(f"📊 Analyzed {len(commits)} commits across {len(file_changes)} files")
        
        return df_commits, df_files

    except git.exc.InvalidGitRepositoryError:
        print("❌ Not a valid git repository.")
        print("Make sure you're in a git repo or provide correct --repo-path")
        return None, None
    except Exception as e:
        print(f"❌ Unexpected error: {e}")
        import traceback
        traceback.print_exc()
        return None, None


def create_html_report(df_commits, df_files, html_path, base_branch, compare_branch):
    """Create a nice self-contained HTML report"""
    try:
        total_commits = len(df_commits)
        total_files = len(df_files)
        net_lines = df_commits['Lines_Added'].sum() - df_commits['Lines_Deleted'].sum()
        
        html_content = f"""<!DOCTYPE html>
<html>
<head>
    <title>Branch Delta Report</title>
    <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
    <style>
        body {{ font-family: Arial, sans-serif; margin: 40px; background: #f8f9fa; }}
        .card {{ background: white; padding: 20px; margin: 20px 0; border-radius: 8px; box-shadow: 0 2px 8px rgba(0,0,0,0.1); }}
        table {{ width: 100%; border-collapse: collapse; }}
        th, td {{ padding: 8px; text-align: left; border-bottom: 1px solid #ddd; }}
        th {{ background: #f1f3f5; }}
    </style>
</head>
<body>
    <h1>🔍 Branch Delta Report</h1>
    <p><strong>{base_branch} vs {compare_branch}</strong> | {total_commits} commits | {total_files} files changed | Net lines: {net_lines:+}</p>
    
    <div class="card">
        <h2>Key Insights</h2>
        <ul>
            <li>Total commits analyzed: {total_commits}</li>
            <li>Files impacted: {total_files}</li>
            <li>JIRA tickets extracted: {df_commits['JIRA_Tickets'].str.count(',').sum() + len(df_commits[df_commits['JIRA_Tickets'] != ''])}</li>
        </ul>
    </div>

    <div class="card">
        <h2>Recent Commits</h2>
        <table>
            <tr><th>SHA</th><th>Author</th><th>Date</th><th>Message</th><th>JIRA</th></tr>
"""
        for _, row in df_commits.head(15).iterrows():
            html_content += f"<tr><td>{row['Commit_SHA']}</td><td>{row['Author']}</td><td>{row['Date']}</td><td>{row['Message']}</td><td>{row['JIRA_Tickets']}</td></tr>\n"
        
        html_content += """        </table>
    </div>
    
    <p><em>Run the tool locally for full interactive Plotly charts.</em></p>
</body>
</html>"""
        
        with open(html_path, 'w', encoding='utf-8') as f:
            f.write(html_content)
    except Exception as e:
        print(f"⚠️ Could not create HTML report: {e}")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description='Git Branch Delta Analyzer')
    parser.add_argument('--repo-path', default='.', help='Path to local git repo')
    parser.add_argument('--github-repo', help='GitHub repo in format owner/repo (e.g. openclaw/openclaw)')
    parser.add_argument('--base', required=True, help='Base branch (e.g. main)')
    parser.add_argument('--compare', required=True, help='Compare branch (e.g. feature/xyz or main@{{7.days.ago}})')
    parser.add_argument('--output', default='delta_reports', help='Output directory')
    args = parser.parse_args()

    # Handle remote repo
    final_repo_path = clone_repo_if_needed(args.github_repo, args.repo_path)
    
    try:
        analyze_branches(final_repo_path, args.base, args.compare, args.output)
    finally:
        # Cleanup temp dir if we cloned
        if args.github_repo and final_repo_path != args.repo_path:
            try:
                shutil.rmtree(final_repo_path, ignore_errors=True)
            except:
                pass
