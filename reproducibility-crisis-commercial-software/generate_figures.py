#!/usr/bin/env python3
"""Generate color figures for the reproducibility crisis study."""

import json
import pandas as pd
import numpy as np
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import seaborn as sns
from collections import Counter
from pathlib import Path

# Style setup
sns.set_theme(style="whitegrid", font_scale=1.1)
plt.rcParams.update({
    'figure.dpi': 300,
    'savefig.dpi': 300,
    'figure.figsize': (10, 6),
    'axes.titlesize': 14,
    'axes.labelsize': 12,
    'font.family': 'sans-serif',
})

OUTPUT_DIR = Path("/home/ubuntu/reproducibility_study/output")
FIG_DIR = OUTPUT_DIR / "figures"
FIG_DIR.mkdir(exist_ok=True)

df = pd.read_csv(OUTPUT_DIR / "extracted_data.csv")

# Color palette
STRATUM_COLORS = {
    'Biomedical_Basic': '#2196F3',
    'Clinical_Medicine': '#F44336',
    'Chemistry_Materials': '#FF9800',
    'Physics_Engineering': '#9C27B0',
    'Social_Behavioral': '#4CAF50',
    'Computational_Science': '#00BCD4',
    'Environmental_Earth': '#795548',
}

STRATUM_LABELS = {
    'Biomedical_Basic': 'Biomedical\n(Basic)',
    'Clinical_Medicine': 'Clinical\nMedicine',
    'Chemistry_Materials': 'Chemistry &\nMaterials',
    'Physics_Engineering': 'Physics &\nEngineering',
    'Social_Behavioral': 'Social &\nBehavioral',
    'Computational_Science': 'Computational\nScience',
    'Environmental_Earth': 'Environmental\n& Earth',
}

# ── Figure 1: Software detection rates by stratum ──────────────────────
fig, ax = plt.subplots(figsize=(12, 6))
strata_order = list(STRATUM_LABELS.keys())
x = np.arange(len(strata_order))
width = 0.25

any_sw_rates = []
comm_rates = []
os_rates = []
for s in strata_order:
    sdf = df[df['stratum'] == s]
    any_sw_rates.append((sdf['software_count'] > 0).mean() * 100)
    comm_rates.append(sdf['has_commercial_software'].mean() * 100)
    os_rates.append(sdf['has_opensource_software'].mean() * 100)

bars1 = ax.bar(x - width, any_sw_rates, width, label='Any Software', color='#1976D2', alpha=0.9)
bars2 = ax.bar(x, comm_rates, width, label='Commercial Software', color='#D32F2F', alpha=0.9)
bars3 = ax.bar(x + width, os_rates, width, label='Open-Source Software', color='#388E3C', alpha=0.9)

ax.set_ylabel('Percentage of Papers (%)')
N_total = len(df)
ax.set_title(f'Figure 1. Software Mention Rates by Research Field (N={N_total:,})')
ax.set_xticks(x)
ax.set_xticklabels([STRATUM_LABELS[s] for s in strata_order], fontsize=9)
ax.legend(loc='upper right')
ax.set_ylim(0, max(any_sw_rates) * 1.2)

for bars in [bars1, bars2, bars3]:
    for bar in bars:
        h = bar.get_height()
        if h > 2:
            ax.annotate(f'{h:.0f}%', xy=(bar.get_x() + bar.get_width()/2, h),
                       xytext=(0, 3), textcoords="offset points", ha='center', va='bottom', fontsize=7)

plt.tight_layout()
plt.savefig(FIG_DIR / 'fig1_software_rates_by_field.png', bbox_inches='tight')
plt.close()
print("Figure 1 saved")

# ── Figure 2: Top commercial software frequency ──────────────────────
fig, ax = plt.subplots(figsize=(10, 7))
comm_counter = Counter()
for swlist in df['commercial_software_list'].dropna():
    for sw in str(swlist).split('; '):
        sw = sw.strip()
        if sw and sw != 'nan':
            # Clean up Adobe pattern issue
            if sw.startswith('Adobe \\'):
                sw = 'Adobe (other)'
            comm_counter[sw] += 1

top_commercial = comm_counter.most_common(15)
sw_names = [s[0] for s in top_commercial]
sw_counts = [s[1] for s in top_commercial]

colors_comm = plt.cm.Reds(np.linspace(0.3, 0.9, len(sw_names)))
bars = ax.barh(range(len(sw_names)), sw_counts, color=colors_comm[::-1])
ax.set_yticks(range(len(sw_names)))
ax.set_yticklabels(sw_names)
ax.invert_yaxis()
ax.set_xlabel('Number of Papers')
ax.set_title(f'Figure 2. Most Frequently Used Commercial Software (N={N_total:,} papers)')

for i, (bar, count) in enumerate(zip(bars, sw_counts)):
    ax.text(bar.get_width() + 0.5, bar.get_y() + bar.get_height()/2,
            str(count), ha='left', va='center', fontsize=10, fontweight='bold')

plt.tight_layout()
plt.savefig(FIG_DIR / 'fig2_top_commercial_software.png', bbox_inches='tight')
plt.close()
print("Figure 2 saved")

# ── Figure 3: Version mention rates by field ──────────────────────────
fig, axes = plt.subplots(1, 2, figsize=(14, 6))

# Left: Version mention rate among papers with software
ax = axes[0]
ver_rates = []
for s in strata_order:
    sdf = df[df['stratum'] == s]
    sw_papers = sdf[sdf['software_count'] > 0]
    if len(sw_papers) > 0:
        ver_rates.append(sw_papers['version_mention_rate'].mean() * 100)
    else:
        ver_rates.append(0)

colors = [STRATUM_COLORS[s] for s in strata_order]
bars = ax.bar(range(len(strata_order)), ver_rates, color=colors, alpha=0.85)
ax.set_xticks(range(len(strata_order)))
ax.set_xticklabels([STRATUM_LABELS[s] for s in strata_order], fontsize=8)
ax.set_ylabel('Version Mention Rate (%)')
ax.set_title('(a) Version Mention Rate\n(among papers citing software)')
ax.set_ylim(0, 100)
for bar, rate in zip(bars, ver_rates):
    ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 1,
            f'{rate:.0f}%', ha='center', va='bottom', fontsize=9, fontweight='bold')

# Right: Code and data availability
ax = axes[1]
code_rates = []
data_rates = []
for s in strata_order:
    sdf = df[df['stratum'] == s]
    code_rates.append(sdf['code_available'].mean() * 100)
    data_rates.append(sdf['data_available'].mean() * 100)

x = np.arange(len(strata_order))
width = 0.35
ax.bar(x - width/2, code_rates, width, label='Code Available', color='#1565C0', alpha=0.85)
ax.bar(x + width/2, data_rates, width, label='Data Available', color='#E65100', alpha=0.85)
ax.set_xticks(x)
ax.set_xticklabels([STRATUM_LABELS[s] for s in strata_order], fontsize=8)
ax.set_ylabel('Percentage of Papers (%)')
ax.set_title('(b) Code and Data Availability Statements')
ax.legend()

fig.suptitle('Figure 3. Version Reporting and Data/Code Availability by Research Field', fontsize=13, fontweight='bold', y=1.02)
plt.tight_layout()
plt.savefig(FIG_DIR / 'fig3_version_and_availability.png', bbox_inches='tight')
plt.close()
print("Figure 3 saved")

# ── Figure 4: Estimated replication cost distribution ──────────────────
fig, axes = plt.subplots(1, 2, figsize=(14, 6))

# Left: Cost distribution (non-zero only)
ax = axes[0]
costs_nz = df[df['estimated_replication_cost_usd'] > 0]['estimated_replication_cost_usd']
ax.hist(costs_nz, bins=30, color='#D32F2F', alpha=0.7, edgecolor='black', linewidth=0.5)
ax.axvline(costs_nz.mean(), color='#1565C0', linestyle='--', linewidth=2, label=f'Mean: ${costs_nz.mean():,.0f}')
ax.axvline(costs_nz.median(), color='#FF8F00', linestyle='--', linewidth=2, label=f'Median: ${costs_nz.median():,.0f}')
ax.set_xlabel('Estimated Replication Cost (USD)')
ax.set_ylabel('Number of Papers')
ax.set_title('(a) Cost Distribution\n(papers with commercial software)')
ax.legend()

# Right: Mean cost by stratum
ax = axes[1]
mean_costs = []
for s in strata_order:
    sdf = df[df['stratum'] == s]
    nz = sdf[sdf['estimated_replication_cost_usd'] > 0]
    mean_costs.append(nz['estimated_replication_cost_usd'].mean() if len(nz) > 0 else 0)

colors = [STRATUM_COLORS[s] for s in strata_order]
bars = ax.bar(range(len(strata_order)), mean_costs, color=colors, alpha=0.85)
ax.set_xticks(range(len(strata_order)))
ax.set_xticklabels([STRATUM_LABELS[s] for s in strata_order], fontsize=8)
ax.set_ylabel('Mean Replication Cost (USD)')
ax.set_title('(b) Mean Cost by Research Field\n(papers with commercial software)')
for bar, cost in zip(bars, mean_costs):
    if cost > 0:
        ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 20,
                f'${cost:,.0f}', ha='center', va='bottom', fontsize=9, fontweight='bold')

fig.suptitle('Figure 4. Estimated Software Replication Costs', fontsize=13, fontweight='bold', y=1.02)
plt.tight_layout()
plt.savefig(FIG_DIR / 'fig4_replication_costs.png', bbox_inches='tight')
plt.close()
print("Figure 4 saved")

# ── Figure 5: Commercial vs Open-Source software landscape ────────────
fig, ax = plt.subplots(figsize=(10, 8))

# All software frequency
all_sw_counter = Counter()
for swlist in df['software_mentioned'].dropna():
    for sw in str(swlist).split('; '):
        sw = sw.strip()
        if sw and sw != 'nan':
            if sw.startswith('Adobe \\'):
                sw = 'Adobe (other)'
            all_sw_counter[sw] += 1

# Determine license type for each
from pubmed_sampler import SOFTWARE_PATTERNS
sw_license = {}
for _, name, lt in SOFTWARE_PATTERNS:
    sw_license[name] = lt
# Fix Adobe pattern
sw_license['Adobe (other)'] = 'commercial'

top20 = all_sw_counter.most_common(20)
names = [s[0] for s in top20]
counts = [s[1] for s in top20]
colors = ['#D32F2F' if sw_license.get(n, 'unknown') == 'commercial' else '#388E3C' for n in names]

bars = ax.barh(range(len(names)), counts, color=colors, alpha=0.85)
ax.set_yticks(range(len(names)))
ax.set_yticklabels(names)
ax.invert_yaxis()
ax.set_xlabel('Number of Papers')
ax.set_title(f'Figure 5. Top 20 Software Tools in Published Research (N={N_total:,})')

# Legend
comm_patch = mpatches.Patch(color='#D32F2F', alpha=0.85, label='Commercial')
os_patch = mpatches.Patch(color='#388E3C', alpha=0.85, label='Open-Source')
ax.legend(handles=[comm_patch, os_patch], loc='lower right')

for bar, count in zip(bars, counts):
    ax.text(bar.get_width() + 0.3, bar.get_y() + bar.get_height()/2,
            str(count), ha='left', va='center', fontsize=9, fontweight='bold')

plt.tight_layout()
plt.savefig(FIG_DIR / 'fig5_software_landscape.png', bbox_inches='tight')
plt.close()
print("Figure 5 saved")

# ── Figure 6: Version availability assessment ─────────────────────────
fig, ax = plt.subplots(figsize=(10, 8))

# Count version availability statuses for commercial software
avail_counts = {'available': 0, 'current': 0, 'legacy_available': 0, 'likely_unavailable': 0, 'unknown': 0}
total_commercial_with_version = 0
for va_str in df['version_availability'].dropna():
    if not va_str or va_str == '':
        continue
    try:
        va = json.loads(va_str)
        for sw, status in va.items():
            if sw_license.get(sw, 'unknown') == 'commercial':
                avail_counts[status] = avail_counts.get(status, 0) + 1
                total_commercial_with_version += 1
    except:
        pass

labels = ['Currently Available', 'Legacy Available', 'Likely Unavailable', 'Unknown']
values = [avail_counts.get('current', 0), avail_counts.get('legacy_available', 0),
          avail_counts.get('likely_unavailable', 0), avail_counts.get('unknown', 0)]
colors_avail = ['#4CAF50', '#FFC107', '#F44336', '#9E9E9E']

# Filter out zero-value slices to avoid clutter
filtered = [(l, v, c) for l, v, c in zip(labels, values, colors_avail) if v > 0]
if filtered:
    labels_f, values_f, colors_f = zip(*filtered)
else:
    labels_f, values_f, colors_f = labels, values, colors_avail

# Use external labels with leader lines to avoid text overlap
# Only show autopct for slices large enough to fit text
def autopct_func(pct):
    return f'{pct:.1f}%' if pct >= 5 else ''

wedges, texts, autotexts = ax.pie(
    values_f, colors=colors_f,
    autopct=autopct_func, startangle=90,
    pctdistance=0.75,
    textprops={'fontsize': 12},
    wedgeprops={'linewidth': 1.5, 'edgecolor': 'white'},
)
for t in autotexts:
    t.set_fontweight('bold')
    t.set_fontsize(12)

# Add external legend with counts and percentages
total_v = sum(values_f)
legend_labels = []
for l, v in zip(labels_f, values_f):
    pct = v / total_v * 100 if total_v > 0 else 0
    legend_labels.append(f'{l}: {v:,} ({pct:.1f}%)')
ax.legend(wedges, legend_labels,
          title='Version Status', loc='center left', bbox_to_anchor=(1.0, 0.5),
          fontsize=11, title_fontsize=12, frameon=True, fancybox=True, shadow=True)

ax.set_title(f'Figure 6. Version Availability of Commercial Software\n(n={total_commercial_with_version} software-version pairs)', fontsize=13, pad=20)

plt.tight_layout()
plt.savefig(FIG_DIR / 'fig6_version_availability.png', bbox_inches='tight')
plt.close()
print("Figure 6 saved")

# ── Figure 7: Heatmap - Software usage across fields ──────────────────
fig, ax = plt.subplots(figsize=(14, 8))

# Top 15 software x 7 strata
top15_sw = [s[0] for s in all_sw_counter.most_common(15)]
heatmap_data = []
for s in strata_order:
    sdf = df[df['stratum'] == s]
    row = []
    for sw in top15_sw:
        count = 0
        for swlist in sdf['software_mentioned'].dropna():
            if sw in str(swlist).split('; '):
                count += 1
        row.append(count)
    heatmap_data.append(row)

hm_df = pd.DataFrame(heatmap_data, 
                       index=[STRATUM_LABELS[s].replace('\n', ' ') for s in strata_order],
                       columns=top15_sw)

sns.heatmap(hm_df, annot=True, fmt='d', cmap='YlOrRd', ax=ax, 
            linewidths=0.5, cbar_kws={'label': 'Number of Papers'})
ax.set_title('Figure 7. Software Usage Across Research Fields (Top 15 Software)', fontsize=13)
ax.set_xlabel('Software')
ax.set_ylabel('Research Field')

plt.tight_layout()
plt.savefig(FIG_DIR / 'fig7_software_heatmap.png', bbox_inches='tight')
plt.close()
print("Figure 7 saved")

# ── Figure 8: PMC full-text availability and its impact ───────────────
fig, axes = plt.subplots(1, 2, figsize=(12, 5))

# Left: PMC fulltext availability by stratum
ax = axes[0]
pmc_rates = []
for s in strata_order:
    sdf = df[df['stratum'] == s]
    pmc_rates.append(sdf['has_pmc_fulltext'].mean() * 100)

colors = [STRATUM_COLORS[s] for s in strata_order]
bars = ax.bar(range(len(strata_order)), pmc_rates, color=colors, alpha=0.85)
ax.set_xticks(range(len(strata_order)))
ax.set_xticklabels([STRATUM_LABELS[s] for s in strata_order], fontsize=8)
ax.set_ylabel('PMC Full-Text Available (%)')
ax.set_title('(a) PMC Full-Text Availability')
ax.set_ylim(0, 100)

# Right: Software detection rate WITH vs WITHOUT PMC
ax = axes[1]
with_pmc = df[df['has_pmc_fulltext'] == True]
without_pmc = df[df['has_pmc_fulltext'] == False]
categories = ['Any Software', 'Commercial\nSoftware', 'Open-Source\nSoftware']
with_rates = [
    (with_pmc['software_count'] > 0).mean() * 100,
    with_pmc['has_commercial_software'].mean() * 100,
    with_pmc['has_opensource_software'].mean() * 100,
]
without_rates = [
    (without_pmc['software_count'] > 0).mean() * 100,
    without_pmc['has_commercial_software'].mean() * 100,
    without_pmc['has_opensource_software'].mean() * 100,
]
x = np.arange(len(categories))
width = 0.35
ax.bar(x - width/2, with_rates, width, label='With PMC Full-Text', color='#1976D2', alpha=0.85)
ax.bar(x + width/2, without_rates, width, label='Abstract Only', color='#B0BEC5', alpha=0.85)
ax.set_xticks(x)
ax.set_xticklabels(categories)
ax.set_ylabel('Detection Rate (%)')
ax.set_title('(b) Software Detection: Full-Text vs Abstract')
ax.legend()

fig.suptitle('Figure 8. Impact of Full-Text Access on Software Detection', fontsize=13, fontweight='bold', y=1.02)
plt.tight_layout()
plt.savefig(FIG_DIR / 'fig8_pmc_impact.png', bbox_inches='tight')
plt.close()
print("Figure 8 saved")

print("\nAll figures generated successfully!")
print(f"Saved to: {FIG_DIR}")
