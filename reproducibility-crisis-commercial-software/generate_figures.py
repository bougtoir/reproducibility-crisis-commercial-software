#!/usr/bin/env python3
"""Generate color figures for the reproducibility crisis study (R1 revision).

Changes from original:
- Combined Fig 2+3 into a single 2-panel figure (Reviewer 1)
- Added "Total" row to heatmap (Reviewer 1)
- Improved Fig 2 title precision (Reviewer 1)
- Fixed Fig 3 internal title (was "Figure 3", now correctly matches paper numbering)
- Added Fig 7: Country/income-group analysis (Reviewer 1)
- Added Fig 8: PMC-only subanalysis (Reviewer 1)
- Fixed OUTPUT_DIR to use relative paths
"""

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

SCRIPT_DIR = Path(__file__).resolve().parent
OUTPUT_DIR = SCRIPT_DIR / "output"
FIG_DIR = OUTPUT_DIR / "figures"
FIG_DIR.mkdir(parents=True, exist_ok=True)

df = pd.read_csv(OUTPUT_DIR / "extracted_data.csv")
N_total = len(df)

# --- Software counters ---
from pubmed_sampler import SOFTWARE_PATTERNS
sw_license = {}
for _, name, lt in SOFTWARE_PATTERNS:
    sw_license[name] = lt
sw_license['Adobe (other)'] = 'commercial'

all_sw_counter = Counter()
for swlist in df['software_mentioned'].dropna():
    for sw in str(swlist).split('; '):
        sw = sw.strip()
        if sw and sw != 'nan':
            if sw.startswith('Adobe \\'):
                sw = 'Adobe (other)'
            all_sw_counter[sw] += 1

comm_counter = Counter()
for swlist in df['commercial_software_list'].dropna():
    for sw in str(swlist).split('; '):
        sw = sw.strip()
        if sw and sw != 'nan':
            if sw.startswith('Adobe \\'):
                sw = 'Adobe (other)'
            comm_counter[sw] += 1

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
strata_order = list(STRATUM_LABELS.keys())

# =====================================================================
# Figure 1: Software detection rates by stratum
# =====================================================================
fig, ax = plt.subplots(figsize=(12, 6))
x = np.arange(len(strata_order))
width = 0.25

any_sw_rates, comm_rates, os_rates = [], [], []
for s in strata_order:
    sdf = df[df['stratum'] == s]
    any_sw_rates.append((sdf['software_count'] > 0).mean() * 100)
    comm_rates.append(sdf['has_commercial_software'].mean() * 100)
    os_rates.append(sdf['has_opensource_software'].mean() * 100)

bars1 = ax.bar(x - width, any_sw_rates, width, label='Any Software', color='#1976D2', alpha=0.9)
bars2 = ax.bar(x, comm_rates, width, label='Commercial Software', color='#D32F2F', alpha=0.9)
bars3 = ax.bar(x + width, os_rates, width, label='Open-Source Software', color='#388E3C', alpha=0.9)

ax.set_ylabel('Percentage of Papers (%)')
ax.set_title(f'Figure 1. Software Mention Rates by Research Field (N = {N_total:,} articles, 2020\u20132026)')
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

# =====================================================================
# Figure 2 (COMBINED): Top 20 software landscape + usage heatmap
# Reviewer 1: combine old Fig 2+3; add "Total" row; improve title
# =====================================================================
fig, axes = plt.subplots(1, 2, figsize=(20, 9), gridspec_kw={'width_ratios': [1, 1.3]})

# Panel (a): Top 20 software tools (horizontal bar, colored by license)
ax = axes[0]
top20 = all_sw_counter.most_common(20)
names = [s[0] for s in top20]
counts = [s[1] for s in top20]
colors_bar = ['#D32F2F' if sw_license.get(n, 'unknown') == 'commercial' else '#388E3C' for n in names]

bars = ax.barh(range(len(names)), counts, color=colors_bar, alpha=0.85)
ax.set_yticks(range(len(names)))
ax.set_yticklabels(names)
ax.invert_yaxis()
ax.set_xlabel('Number of Articles')
ax.set_title('(a) Top 20 Software Declared in Published\nResearch Articles (2020\u20132026)')

comm_patch = mpatches.Patch(color='#D32F2F', alpha=0.85, label='Commercial')
os_patch = mpatches.Patch(color='#388E3C', alpha=0.85, label='Open-Source')
ax.legend(handles=[comm_patch, os_patch], loc='lower right')

for bar, count in zip(bars, counts):
    ax.text(bar.get_width() + 0.3, bar.get_y() + bar.get_height()/2,
            str(count), ha='left', va='center', fontsize=9, fontweight='bold')

# Panel (b): Heatmap of top 15 software across fields + Total row
ax = axes[1]
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

# Add "Total" row (Reviewer 1 request)
total_row = [sum(col) for col in zip(*heatmap_data)]
heatmap_data.append(total_row)
row_labels = [STRATUM_LABELS[s].replace('\n', ' ') for s in strata_order] + ['Total']

hm_df = pd.DataFrame(heatmap_data, index=row_labels, columns=top15_sw)

sns.heatmap(hm_df, annot=True, fmt='d', cmap='YlOrRd', ax=ax,
            linewidths=0.5, cbar_kws={'label': 'Number of Articles'})
ax.set_title('(b) Software Usage Across Research Fields')
ax.set_xlabel('Software')
ax.set_ylabel('')

fig.suptitle(f'Figure 2. Software Landscape in Published Research (N = {N_total:,} articles)',
             fontsize=14, fontweight='bold', y=1.01)
plt.tight_layout()
plt.savefig(FIG_DIR / 'fig2_software_landscape_combined.png', bbox_inches='tight')
plt.close()
print("Figure 2 (combined) saved")

# =====================================================================
# Figure 3: Version mention rates + code/data availability
# Fixed internal title (Reviewer 1: was "Figure 3" while paper called it Fig 4)
# =====================================================================
fig, axes = plt.subplots(1, 2, figsize=(14, 6))

# (a) Version mention rate among papers with software
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
ax.set_title('(a) Version Mention Rate\n(among articles reporting software)')
ax.set_ylim(0, 100)
for bar, rate in zip(bars, ver_rates):
    ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 1,
            f'{rate:.0f}%', ha='center', va='bottom', fontsize=9, fontweight='bold')

# (b) Code and data availability
ax = axes[1]
code_rates, data_rates = [], []
for s in strata_order:
    sdf = df[df['stratum'] == s]
    code_rates.append(sdf['code_available'].mean() * 100)
    data_rates.append(sdf['data_available'].mean() * 100)

x = np.arange(len(strata_order))
w = 0.35
ax.bar(x - w/2, code_rates, w, label='Code Available', color='#1565C0', alpha=0.85)
ax.bar(x + w/2, data_rates, w, label='Data Available', color='#E65100', alpha=0.85)
ax.set_xticks(x)
ax.set_xticklabels([STRATUM_LABELS[s] for s in strata_order], fontsize=8)
ax.set_ylabel('Percentage of Articles (%)')
ax.set_title('(b) Code and Data Availability Statements')
ax.legend()

fig.suptitle('Figure 3. Version Reporting and Data/Code Availability by Research Field',
             fontsize=13, fontweight='bold', y=1.02)
plt.tight_layout()
plt.savefig(FIG_DIR / 'fig3_version_and_availability.png', bbox_inches='tight')
plt.close()
print("Figure 3 saved")

# =====================================================================
# Figure 4: Version availability assessment (pie chart)
# =====================================================================
fig, ax = plt.subplots(figsize=(10, 8))

avail_counts = {'available': 0, 'current': 0, 'legacy_available': 0,
                'likely_unavailable': 0, 'unknown': 0}
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
    except Exception:
        pass

labels = ['Currently Available', 'Legacy Available', 'Likely Unavailable', 'Unknown']
values = [avail_counts.get('current', 0), avail_counts.get('legacy_available', 0),
          avail_counts.get('likely_unavailable', 0), avail_counts.get('unknown', 0)]
colors_avail = ['#4CAF50', '#FFC107', '#F44336', '#9E9E9E']

filtered = [(l, v, c) for l, v, c in zip(labels, values, colors_avail) if v > 0]
if filtered:
    labels_f, values_f, colors_f = zip(*filtered)
else:
    labels_f, values_f, colors_f = labels, values, colors_avail

def autopct_func(pct):
    return f'{pct:.1f}%' if pct >= 5 else ''

wedges, texts, autotexts = ax.pie(
    values_f, colors=colors_f, autopct=autopct_func, startangle=90,
    pctdistance=0.75, textprops={'fontsize': 12},
    wedgeprops={'linewidth': 1.5, 'edgecolor': 'white'},
)
for t in autotexts:
    t.set_fontweight('bold')
    t.set_fontsize(12)

total_v = sum(values_f)
legend_labels = []
for l, v in zip(labels_f, values_f):
    pct = v / total_v * 100 if total_v > 0 else 0
    legend_labels.append(f'{l}: {v:,} ({pct:.1f}%)')
ax.legend(wedges, legend_labels, title='Version Status', loc='center left',
          bbox_to_anchor=(1.0, 0.5), fontsize=11, title_fontsize=12,
          frameon=True, fancybox=True, shadow=True)
ax.set_title(f'Figure 4. Version Availability of Commercial Software\n'
             f'(n = {total_commercial_with_version} software\u2013version pairs)',
             fontsize=13, pad=20)

plt.tight_layout()
plt.savefig(FIG_DIR / 'fig4_version_availability.png', bbox_inches='tight')
plt.close()
print("Figure 4 saved")

# =====================================================================
# Figure 5: Estimated replication cost distribution
# =====================================================================
fig, axes = plt.subplots(1, 2, figsize=(14, 6))

costs_nz = df[df['estimated_replication_cost_usd'] > 0]['estimated_replication_cost_usd']

ax = axes[0]
ax.hist(costs_nz, bins=30, color='#D32F2F', alpha=0.7, edgecolor='black', linewidth=0.5)
ax.axvline(costs_nz.mean(), color='#1565C0', linestyle='--', linewidth=2,
           label=f'Mean: ${costs_nz.mean():,.0f}')
ax.axvline(costs_nz.median(), color='#FF8F00', linestyle='--', linewidth=2,
           label=f'Median: ${costs_nz.median():,.0f}')
ax.set_xlabel('Estimated Replication Cost (USD)')
ax.set_ylabel('Number of Articles')
ax.set_title('(a) Cost Distribution\n(articles with commercial software)')
ax.legend()

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
ax.set_title('(b) Mean Cost by Research Field\n(articles with commercial software)')
for bar, cost in zip(bars, mean_costs):
    if cost > 0:
        ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 20,
                f'${cost:,.0f}', ha='center', va='bottom', fontsize=9, fontweight='bold')

fig.suptitle('Figure 5. Estimated Software Replication Costs', fontsize=13, fontweight='bold', y=1.02)
plt.tight_layout()
plt.savefig(FIG_DIR / 'fig5_replication_costs.png', bbox_inches='tight')
plt.close()
print("Figure 5 saved")

# =====================================================================
# Figure 6: PMC full-text availability and its impact
# =====================================================================
fig, axes = plt.subplots(1, 2, figsize=(12, 5))

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
w = 0.35
ax.bar(x - w/2, with_rates, w, label='With PMC Full-Text', color='#1976D2', alpha=0.85)
ax.bar(x + w/2, without_rates, w, label='Abstract Only', color='#B0BEC5', alpha=0.85)
ax.set_xticks(x)
ax.set_xticklabels(categories)
ax.set_ylabel('Detection Rate (%)')
ax.set_title('(b) Software Detection: Full-Text vs Abstract')
ax.legend()

fig.suptitle('Figure 6. Impact of Full-Text Access on Software Detection',
             fontsize=13, fontweight='bold', y=1.02)
plt.tight_layout()
plt.savefig(FIG_DIR / 'fig6_pmc_impact.png', bbox_inches='tight')
plt.close()
print("Figure 6 saved")

# =====================================================================
# Figure 7 (NEW): Country / income-group analysis (Reviewer 1)
# =====================================================================
country_map = {
    'USA': 'United States', 'United States of America': 'United States',
    'China': 'China', 'PR China': 'China', "People's Republic of China": 'China',
    'UK': 'United Kingdom', 'Republic of Korea': 'South Korea', 'Korea': 'South Korea',
}
df['country_norm'] = df['country'].map(lambda x: country_map.get(x, x) if pd.notna(x) else x)

hic_countries = {
    'United States', 'Japan', 'Germany', 'Italy', 'United Kingdom', 'Canada',
    'Australia', 'France', 'Spain', 'South Korea', 'Netherlands', 'Sweden',
    'Switzerland', 'Belgium', 'Austria', 'Denmark', 'Finland', 'Norway',
    'Ireland', 'Israel', 'Singapore', 'New Zealand', 'Portugal', 'Greece',
    'Czech Republic', 'Poland', 'Saudi Arabia', 'Chile', 'Hungary', 'Croatia',
    'Slovakia', 'Slovenia', 'Lithuania', 'Latvia', 'Estonia', 'Luxembourg',
    'Iceland', 'Cyprus', 'Malta', 'Taiwan', 'Hong Kong', 'Qatar', 'UAE',
    'Kuwait', 'Oman', 'Bahrain', 'Romania', 'Bulgaria', 'Uruguay', 'Panama',
    'Puerto Rico',
}
df['income_group'] = df['country_norm'].apply(
    lambda x: 'HIC' if x in hic_countries else ('LMIC' if pd.notna(x) else None)
)

fig, axes = plt.subplots(1, 3, figsize=(18, 6))

# (a) Commercial sw rate among sw-using papers by income group
ax = axes[0]
groups = ['HIC', 'LMIC']
comm_by_income = []
for g in groups:
    sub = df[(df['income_group'] == g) & (df['software_count'] > 0)]
    comm_by_income.append(sub['has_commercial_software'].mean() * 100)
bars = ax.bar(groups, comm_by_income, color=['#1976D2', '#D32F2F'], alpha=0.85, width=0.5)
ax.set_ylabel('Commercial Software Use (%)')
ax.set_title('(a) Commercial Software Use\n(among articles reporting software)')
for bar, val in zip(bars, comm_by_income):
    ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 0.5,
            f'{val:.1f}%', ha='center', va='bottom', fontsize=11, fontweight='bold')

# (b) Mean replication cost by income group
ax = axes[1]
cost_by_income = []
for g in groups:
    sub = df[(df['income_group'] == g) & (df['estimated_replication_cost_usd'] > 0)]
    cost_by_income.append(sub['estimated_replication_cost_usd'].mean())
bars = ax.bar(groups, cost_by_income, color=['#1976D2', '#D32F2F'], alpha=0.85, width=0.5)
ax.set_ylabel('Mean Replication Cost (USD)')
ax.set_title('(b) Mean Replication Cost\n(articles with commercial software)')
for bar, val in zip(bars, cost_by_income):
    ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 20,
            f'${val:,.0f}', ha='center', va='bottom', fontsize=11, fontweight='bold')

# (c) Code availability by income group
ax = axes[2]
code_by_income = []
for g in groups:
    sub = df[df['income_group'] == g]
    code_by_income.append(sub['code_available'].mean() * 100)
bars = ax.bar(groups, code_by_income, color=['#1976D2', '#D32F2F'], alpha=0.85, width=0.5)
ax.set_ylabel('Code Available (%)')
ax.set_title('(c) Code Availability\nStatements')
for bar, val in zip(bars, code_by_income):
    ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 0.2,
            f'{val:.1f}%', ha='center', va='bottom', fontsize=11, fontweight='bold')

fig.suptitle('Figure 8. Software Use and Reproducibility Indicators by Country Income Group\n'
             '(World Bank classification: HIC = High-Income Countries, LMIC = Low- and Middle-Income Countries)',
             fontsize=13, fontweight='bold', y=1.04)
plt.tight_layout()
plt.savefig(FIG_DIR / 'fig8_country_income.png', bbox_inches='tight')
plt.close()
print("Figure 8 (country/income) saved")

# =====================================================================
# Figure 8 (NEW): PMC-only subanalysis (Reviewer 1 suggestion)
# =====================================================================
pmc_df = df[df['has_pmc_fulltext'] == True]
pmc_sw = pmc_df[pmc_df['software_count'] > 0]

fig, axes = plt.subplots(1, 2, figsize=(14, 6))

# (a) Software detection rates for PMC-only subset, by stratum
ax = axes[0]
pmc_any, pmc_comm, pmc_os = [], [], []
for s in strata_order:
    sdf = pmc_df[pmc_df['stratum'] == s]
    pmc_any.append((sdf['software_count'] > 0).mean() * 100)
    pmc_comm.append(sdf['has_commercial_software'].mean() * 100)
    pmc_os.append(sdf['has_opensource_software'].mean() * 100)

x = np.arange(len(strata_order))
w = 0.25
ax.bar(x - w, pmc_any, w, label='Any Software', color='#1976D2', alpha=0.9)
ax.bar(x, pmc_comm, w, label='Commercial Software', color='#D32F2F', alpha=0.9)
ax.bar(x + w, pmc_os, w, label='Open-Source Software', color='#388E3C', alpha=0.9)
ax.set_xticks(x)
ax.set_xticklabels([STRATUM_LABELS[s] for s in strata_order], fontsize=8)
ax.set_ylabel('Percentage of Articles (%)')
ax.set_title(f'(a) Software Detection (PMC Full-Text Only, N = {len(pmc_df):,})')
ax.legend(loc='upper right', fontsize=8)
ax.set_ylim(0, max(pmc_any) * 1.25)

# (b) Commercial sw rate among sw-users: full sample vs PMC-only
ax = axes[1]
full_sw = df[df['software_count'] > 0]
categories = ['Full Sample', 'PMC Full-Text\nSubset']
comm_among_sw = [
    full_sw['has_commercial_software'].mean() * 100,
    pmc_sw['has_commercial_software'].mean() * 100,
]
ver_among_sw = [
    full_sw['version_mention_rate'].mean() * 100,
    pmc_sw['version_mention_rate'].mean() * 100,
]
x = np.arange(len(categories))
w = 0.3
b1 = ax.bar(x - w/2, comm_among_sw, w, label='Commercial SW (%)', color='#D32F2F', alpha=0.85)
b2 = ax.bar(x + w/2, ver_among_sw, w, label='Version Reported (%)', color='#FF8F00', alpha=0.85)
ax.set_xticks(x)
ax.set_xticklabels(categories)
ax.set_ylabel('Percentage (%)')
ax.set_title('(b) Comparison: Full Sample vs PMC Subset\n(among articles reporting software)')
ax.legend()
for bar_set in [b1, b2]:
    for bar in bar_set:
        h = bar.get_height()
        ax.text(bar.get_x() + bar.get_width()/2, h + 0.5,
                f'{h:.1f}%', ha='center', va='bottom', fontsize=10, fontweight='bold')

fig.suptitle(f'Figure 7. Sensitivity Analysis: PMC Full-Text Subset (N = {len(pmc_df):,})',
             fontsize=13, fontweight='bold', y=1.02)
plt.tight_layout()
plt.savefig(FIG_DIR / 'fig7_pmc_subanalysis.png', bbox_inches='tight')
plt.close()
print("Figure 7 (PMC subanalysis) saved")

print("\nAll figures generated successfully!")
print(f"Saved to: {FIG_DIR}")
