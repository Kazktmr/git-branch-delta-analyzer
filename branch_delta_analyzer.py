# !/usr/bin/env python3
"""
Git Branch Delta Analyzer for TPMs
Compares two branches (local or remote GitHub) and generates Excel + HTML reports.

Now with significantly improved interactive HTML dashboard.
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
except ImportError as e:
    print(f"❌ Missing dependency: {e}")
    print("Run: pip install gitpython pandas openpyxl")
    sys.exit(1)

try:
    from plotly.subplots import make_subplots
    import plotly.graph_objects as go
except ImportError:
    print("⚠️ Plotly not installed - HTML charts will use Chart.js fallback")


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
        
        commits = list(repo.iter_commits(f'{base_branch}..{compare_branch}'))
        
        if not commits:
            print("⚠️ No differences found between branches.")
            return None, None

        data = []
        file_changes = {}

        for commit in commits:
            jiras = parse_jira_tickets(commit.message)
            jira_str = ', '.join(jiras) if jiras else ''
            
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

        os.makedirs(output_dir, exist_ok=True)
        timestamp = datetime.now().strftime('%Y%m%d_%H%M')
        
        excel_path = os.path.join(output_dir, f'delta_summary_{timestamp}.xlsx')
        with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
            df_commits.to_excel(writer, sheet_name='Commits', index=False)
            df_files.to_excel(writer, sheet_name='Files_Changed', index=False)
        
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
    """Create a beautiful, highly interactive HTML dashboard with nice colors and digestible layout for TPMs"""
    try:
        total_commits = len(df_commits)
        total_files = len(df_files)
        total_added = int(df_commits['Lines_Added'].sum())
        total_deleted = int(df_commits['Lines_Deleted'].sum())
        net_lines = total_added - total_deleted
        
        # Prepare data for charts
        top_authors = df_commits['Author'].value_counts().head(8).to_dict()
        author_labels = list(top_authors.keys())
        author_values = list(top_authors.values())
        
        top_files = df_files.head(10)
        file_labels = list(top_files['File_Path'])
        file_values = list(top_files['Commits_Touching'])
        
        # Simple timeline data (group by date)
        df_commits['DateOnly'] = pd.to_datetime(df_commits['Date']).dt.date
        timeline = df_commits.groupby('DateOnly').size().reset_index(name='count')
        timeline_labels = [str(d) for d in timeline['DateOnly']]
        timeline_values = list(timeline['count'])
        
        html_content = f"""<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Branch Delta Report • {base_branch} vs {compare_branch}</title>
    <script src="https://cdn.tailwindcss.com"></script>
    <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600&display=swap');
        body {{ font-family: 'Inter', system-ui, sans-serif; }}
        .kpi-value {{ font-size: 2.75rem; line-height: 1; }}
    </style>
</head>
<body class="bg-slate-50">
    <div class="max-w-screen-2xl mx-auto p-8">
        <!-- Header -->
        <div class="flex justify-between items-start mb-10">
            <div>
                <div class="flex items-center gap-x-3">
                    <div class="text-5xl">🔍</div>
                    <h1 class="text-4xl font-semibold tracking-tight text-slate-900">Branch Delta Report</h1>
                </div>
                <p class="text-xl text-slate-600 mt-2">{base_branch} <span class="text-indigo-500 font-medium">vs</span> {compare_branch}</p>
            </div>
            <div class="text-right">
                <span class="inline-flex items-center px-5 py-2 bg-emerald-100 text-emerald-700 rounded-3xl text-sm font-semibold">Generated {datetime.now().strftime('%B %d, %Y at %H:%M')}</span>
            </div>
        </div>

        <!-- KPI Cards -->
        <div class="grid grid-cols-2 lg:grid-cols-4 gap-6 mb-12">
            <div class="bg-white rounded-3xl shadow p-8">
                <div class="text-slate-500 text-sm font-medium">COMMITS</div>
                <div class="kpi-value text-slate-900 font-semibold mt-3">{total_commits}</div>
                <div class="text-emerald-600 text-sm mt-1">in delta</div>
            </div>
            <div class="bg-white rounded-3xl shadow p-8">
                <div class="text-slate-500 text-sm font-medium">FILES CHANGED</div>
                <div class="kpi-value text-slate-900 font-semibold mt-3">{total_files}</div>
            </div>
            <div class="bg-white rounded-3xl shadow p-8">
                <div class="text-slate-500 text-sm font-medium">LINES ADDED</div>
                <div class="kpi-value text-emerald-600 font-semibold mt-3">+{total_added}</div>
            </div>
            <div class="bg-white rounded-3xl shadow p-8">
                <div class="text-slate-500 text-sm font-medium">NET CHANGE</div>
                <div class="kpi-value {'text-emerald-600' if net_lines >= 0 else 'text-rose-600'} font-semibold mt-3">{net_lines:+,}</div>
            </div>
        </div>

        <div class="grid grid-cols-1 lg:grid-cols-12 gap-8">
            <!-- Timeline -->
            <div class="lg:col-span-7 bg-white rounded-3xl shadow p-8">
                <h3 class="text-lg font-semibold mb-6 text-slate-700">Commit Timeline</h3>
                <canvas id="timeline" class="w-full"></canvas>
            </div>

            <!-- Top Contributors -->
            <div class="lg:col-span-5 bg-white rounded-3xl shadow p-8">
                <h3 class="text-lg font-semibold mb-6 text-slate-700">Top Contributors</h3>
                <canvas id="contributors" class="w-full"></canvas>
            </div>

            <!-- Top Files -->
            <div class="lg:col-span-12 bg-white rounded-3xl shadow p-8">
                <h3 class="text-lg font-semibold mb-6 text-slate-700">Most Changed Files</h3>
                <canvas id="files" class="w-full"></canvas>
            </div>
        </div>

        <!-- Commits Table -->
        <div class="mt-12 bg-white rounded-3xl shadow overflow-hidden">
            <div class="px-8 pt-8 pb-4 flex items-center justify-between">
                <h3 class="text-lg font-semibold text-slate-700">All Commits in Delta</h3>
                <input id="search" type="text" placeholder="Search message or JIRA ticket..." class="bg-slate-100 border-0 rounded-3xl px-5 py-3 text-sm w-80 focus:ring-2 focus:ring-indigo-300">
            </div>
            <div class="overflow-x-auto">
                <table class="w-full" id="commitsTable">
                    <thead>
                        <tr class="bg-slate-50 border-b">
                            <th class="text-left px-8 py-5 font-medium text-slate-600">Date</th>
                            <th class="text-left px-8 py-5 font-medium text-slate-600">Author</th>
                            <th class="text-left px-8 py-5 font-medium text-slate-600">Message</th>
                            <th class="text-left px-8 py-5 font-medium text-slate-600">JIRA Tickets</th>
                            <th class="text-right px-8 py-5 font-medium text-slate-600">Files</th>
                            <th class="text-right px-8 py-5 font-medium text-emerald-600">+ Lines</th>
                            <th class="text-right px-8 py-5 font-medium text-rose-600">- Lines</th>
                        </tr>
                    </thead>
                    <tbody class="divide-y text-sm">
                        {''.join([f'<tr class="hover:bg-slate-50"><td class="px-8 py-4">{row.Date}</td><td class="px-8 py-4 font-medium">{row.Author}</td><td class="px-8 py-4 max-w-md">{row.Message}</td><td class="px-8 py-4 text-indigo-600 font-medium">{row.JIRA_Tickets or "—"}</td><td class="px-8 py-4 text-right font-medium">{row.Files_Changed}</td><td class="px-8 py-4 text-right text-emerald-600 font-medium">{row.Lines_Added}</td><td class="px-8 py-4 text-right text-rose-600 font-medium">{row.Lines_Deleted}</td></tr>' for _, row in df_commits.iterrows()])}
                    </tbody>
                </table>
            </div>
        </div>
    </div>

    <script>
        // Tailwind script
        function initializeTailwind() {{
            return;
        }}

        // Interactive Charts
        document.addEventListener('DOMContentLoaded', function () {{
            // Timeline
            new Chart(document.getElementById('timeline'), {{
                type: 'line',
                data: {{
                    labels: {timeline_labels},
                    datasets: [{{
                        label: 'Commits',
                        data: {timeline_values},
                        borderColor: '#6366f1',
                        backgroundColor: 'rgba(99, 102, 241, 0.1)',
                        tension: 0.3
                    }}]
                }},
                options: {{
                    plugins: {{ legend: {{ display: false }} }},
                    scales: {{ y: {{ beginAtZero: true }} }}
                }}
            }});

            // Contributors
            new Chart(document.getElementById('contributors'), {{
                type: 'doughnut',
                data: {{
                    labels: {author_labels},
                    datasets: [{{
                        data: {author_values},
                        backgroundColor: ['#6366f1', '#22c55e', '#eab308', '#ec4899', '#06b67f', '#a855f7', '#f59e0b', '#ef4444']
                    }}]
                }},
                options: {{ cutout: '70%' }}
            }});

            // Files
            new Chart(document.getElementById('files'), {{
                type: 'bar',
                data: {{
                    labels: {file_labels},
                    datasets: [{{
                        label: 'Commits touching file',
                        data: {file_values},
                        backgroundColor: '#6366f1'
                    }}]
                }},
                options: {{
                    indexAxis: 'y',
                    plugins: {{ legend: {{ display: false }} }}
                }}
            }});

            // Simple search for table
            const searchInput = document.getElementById('search');
            const tableBody = document.querySelector('#commitsTable tbody');
            searchInput.addEventListener('input', function () {{
                const term = this.value.toLowerCase();
                Array.from(tableBody.rows).forEach(row => {{
                    const text = row.textContent.toLowerCase();
                    row.style.display = text.includes(term) ? '' : 'none';
                }});
            }});
        }});
    </script>
</body>
</html>"""
        
        with open(html_path, 'w', encoding='utf-8') as f:
            f.write(html_content)
        print("🎨 Enhanced interactive HTML report created with beautiful colors and TPM-friendly design")
    except Exception as e:
        print(f"⚠️ Could not create HTML report: {e}")
        # Fallback
        with open(html_path, 'w', encoding='utf-8') as f:
            f.write("<h1>Error creating HTML report</h1><p>Please check console for details.</p>")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description='Git Branch Delta Analyzer for TPMs - with beautiful interactive HTML')
    parser.add_argument('--repo-path', default='.', help='Path to local git repo')
    parser.add_argument('--github-repo', help='GitHub repo in format owner/repo (e.g. openclaw/openclaw)')
    parser.add_argument('--base', required=True, help='Base branch (e.g. main)')
    parser.add_argument('--compare', required=True, help='Compare branch (e.g. feature/xyz or main@{7.days.ago})')
    parser.add_argument('--output', default='delta_reports', help='Output directory')
    args = parser.parse_args()

    final_repo_path = clone_repo_if_needed(args.github_repo, args.repo_path)
    
    try:
        analyze_branches(final_repo_path, args.base, args.compare, args.output)
    finally:
        if args.github_repo and final_repo_path != args.repo_path:
            try:
                shutil.rmtree(final_repo_path, ignore_errors=True)
            except:
                pass
"