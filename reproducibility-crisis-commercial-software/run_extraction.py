#!/usr/bin/env python3
"""Resume extraction from already-sampled PMIDs."""
import json
import random
import pandas as pd
from pathlib import Path
from pubmed_sampler import (
    process_papers, check_version_availability, estimate_replication_cost,
    generate_summary, OUTPUT_DIR, logger
)

random.seed(42)

# Load already sampled PMIDs
pmids_df = pd.read_csv(OUTPUT_DIR / "sampled_pmids.csv")
all_pmids = {}
for stratum in pmids_df["stratum"].unique():
    all_pmids[stratum] = pmids_df[pmids_df["stratum"] == stratum]["pmid"].astype(str).tolist()

logger.info(f"Loaded {len(pmids_df)} PMIDs from {len(all_pmids)} strata")

# Load stratum stats
stats_df = pd.read_csv(OUTPUT_DIR / "sampling_stats.csv", index_col=0)
stratum_stats = stats_df.T.to_dict()

# Process papers
all_records = process_papers(all_pmids, stratum_stats)

# Add availability and cost
for record in all_records:
    if record["software_versions"]:
        versions = json.loads(record["software_versions"])
        availability = {}
        for sw, ver in versions.items():
            availability[sw] = check_version_availability(sw, ver)
        record["version_availability"] = json.dumps(availability)
    else:
        record["version_availability"] = ""
    record["estimated_replication_cost_usd"] = estimate_replication_cost(
        record.get("commercial_software_list", "")
    )

# Save
df = pd.DataFrame(all_records)
df.to_csv(OUTPUT_DIR / "extracted_data.csv", index=False)
logger.info(f"\nSaved {len(df)} records to extracted_data.csv")

# Summary
generate_summary(df, stratum_stats)
logger.info("Extraction complete!")
