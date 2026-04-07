# Amazon Ads SKU Migration Tools

Tools for bulk-managing SKUs across Amazon Advertising campaigns.

## The Problem

When you switch barcode types on Amazon, your SKUs change but your advertising campaigns still reference the old ones. Amazon's bulk operations don't let you update a SKU on an existing Product Ad - you have to create new ones.

Doing this manually across dozens of campaigns with hundreds of Product Ads is painful. This script automates it.

## What It Does

1. Reads your Amazon Advertising bulk download file (`.xlsx`)
2. Finds all Product Ad rows across **Sponsored Products** and **Sponsored Display** campaigns
3. Generates upload-ready files with new Product Ad rows using your updated SKUs
4. Outputs both a **full migration file** and a **single-campaign test file**

### What about Sponsored Brands?

Sponsored Brands campaigns reference products by ASIN, not SKU. Since ASINs don't change when you switch barcode types, SB campaigns don't need modification.

## Quick Start

```bash
# Install dependency
pip install openpyxl

# Download your bulk file from Amazon Advertising
# (Campaign Manager > Bulk Operations > Download)

# Run the script
python generate_sku_migration.py your_bulk_download.xlsx --suffix A
```

This generates two files:
- `SKU_Migration_FULL.xlsx` - All new Product Ad rows
- `SKU_Migration_TEST.xlsx` - Single campaign for safe testing

## Usage

```bash
python generate_sku_migration.py <bulk_file.xlsx> [options]
```

### Options

| Flag | Description | Default |
|------|-------------|---------|
| `--suffix` | Suffix to append to each SKU | `A` |
| `--test-campaign` | Campaign ID for the test file | Auto-selects smallest |
| `--output-dir` | Where to save output files | Same as source file |
| `--full-only` | Skip generating the test file | Off |

### Examples

```bash
# Default: append "A" to all SKUs
python generate_sku_migration.py bulk.xlsx

# Custom suffix
python generate_sku_migration.py bulk.xlsx --suffix _V2

# Specify which campaign to test with
python generate_sku_migration.py bulk.xlsx --test-campaign 329423496598982

# Output to a different directory
python generate_sku_migration.py bulk.xlsx --output-dir ./output
```

## How It Works

For each existing Product Ad row in your Sponsored Products and Sponsored Display sheets, the script creates a new row with:

| Column | Value |
|--------|-------|
| Product | `Sponsored Products` or `Sponsored Display` |
| Entity | `Product Ad` |
| Operation | `Create` |
| Campaign ID | Copied from existing row |
| Ad Group ID | Copied from existing row |
| State | `enabled` |
| SKU | Original SKU + your suffix |

Everything else is left blank - Amazon fills in ASIN, metrics, names, etc. on upload.

### Smart Exclusions

The script automatically skips:
- SKUs that already end with your suffix (avoids double-appending)
- FBA error entries (`FBA...Missing` SKUs)
- SKUs where the new variant already exists in the same ad group (avoids duplicates)

## Recommended Workflow

1. **Download** your bulk file from Amazon Advertising
2. **Run the script** to generate the migration files
3. **Upload the TEST file** first (single campaign)
4. **Wait 24-48 hours** and verify the new Product Ad appears correctly
5. **Upload the FULL file** after the test succeeds
6. **Monitor** the bulk upload processing report for any errors
7. **(Later)** Archive/pause old SKU Product Ads once new ones are confirmed

### Common Errors to Watch For

- **"SKU not found"** - The new SKU hasn't been registered in Seller Central inventory yet
- **"Duplicate product ad"** - The script should prevent this, but check if you've run it multiple times

## Requirements

- Python 3.7+
- `openpyxl` library (`pip install openpyxl`)
- An Amazon Advertising bulk download file (`.xlsx`)

## Important Notes

- **SP vs SD column differences**: Sponsored Display uses column F for Ad Group ID (not E like SP) and column P for State (not R). The script handles this automatically.
- **This only creates new ads** - it never modifies or deletes existing ones. Your current ads remain untouched.
- **Don't include your bulk download in version control** - it contains campaign IDs and performance data. The `.gitignore` excludes `.xlsx` files by default.

---

## Script 2: Add SKU to Existing Ad Groups

Sometimes you need to add a specific SKU to campaigns where related SKUs already exist — for example, when a SKU is tied to a running deal and can't be deleted, but needs to be included in all relevant campaigns.

### What It Does

1. Searches your bulk download for ad groups containing specified SKUs
2. Generates upload-ready rows to add a new SKU to those same ad groups
3. Skips ad groups where the new SKU already exists (safe to re-run)

### Usage

```bash
python add_sku_to_adgroups.py <bulk_file.xlsx> [options]
```

### Options

| Flag | Description | Default |
|------|-------------|---------|
| `--search-skus` | SKUs to search for (space-separated) | *(none — must provide)* |
| `--add-sku` | SKU to add to matched ad groups | *(none — must provide)* |
| `--test-campaign` | Campaign ID for the test file | Auto-selects smallest |
| `--output-dir` | Where to save output files | Same as source file |
| `--full-only` | Skip generating the test file | Off |

### Example

```bash
# Find all ad groups containing OLD-SKU-1 or OLD-SKU-2, add NEW-SKU to each
python add_sku_to_adgroups.py bulk.xlsx --search-skus "OLD-SKU-1" "OLD-SKU-2" --add-sku "NEW-SKU"
```

Outputs:
- `Add_SKU_FULL.xlsx` - All new Product Ad rows
- `Add_SKU_TEST.xlsx` - Single campaign for safe testing

### Important

- **Download an unfiltered bulk sheet** — if you filter by impressions or other metrics, campaigns without activity won't be included and will be missed.
- Like the migration script, this only creates new ads — it never modifies or deletes existing ones.

---

## Built With

This tool was built with [Claude Code](https://claude.ai/code) as part of a real Amazon seller SKU migration. Full walkthrough on [Substack](https://substack.com).

## License

MIT - Use it however you want.
