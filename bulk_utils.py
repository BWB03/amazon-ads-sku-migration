"""
Shared utilities for Amazon Advertising bulk file operations.

Common functions for reading bulk download workbooks, looking up
campaigns, deduplicating entries, and writing upload-ready output files.
"""

# Sheets that support negative keywords
NEGATIVE_KEYWORD_SHEETS = [
    "Sponsored Products Campaigns",
    "Sponsored Brands Campaigns",
]

SHEET_PRODUCT_VALUES = {
    "Sponsored Products Campaigns": "Sponsored Products",
    "Sponsored Display Campaigns": "Sponsored Display",
    "Sponsored Brands Campaigns": "Sponsored Brands",
}


def find_columns_by_header(ws):
    """Read the header row and return {header_name: 1-based column index}."""
    header_map = {}
    for cell in next(ws.iter_rows(min_row=1, max_row=1)):
        if cell.value:
            header_map[str(cell.value).strip()] = cell.column
    return header_map


def find_column(header_map, *names):
    """Find a column index by trying multiple possible header names.
    Returns 1-based column index or None.
    """
    for name in names:
        if name in header_map:
            return header_map[name]
    return None


def build_campaign_lookup(wb, sheet_names=None):
    """
    Scan sheets for Campaign and Ad Group entity rows.

    Returns:
        {campaign_name: {campaign_id, sheet_name, ad_groups: {ag_name: ag_id}}}
    """
    if sheet_names is None:
        sheet_names = NEGATIVE_KEYWORD_SHEETS

    campaigns = {}

    for sheet_name in sheet_names:
        if sheet_name not in wb.sheetnames:
            continue

        ws = wb[sheet_name]
        header_map = find_columns_by_header(ws)

        entity_col = find_column(header_map, "Entity")
        campaign_id_col = find_column(header_map, "Campaign ID")
        campaign_name_col = find_column(
            header_map, "Campaign Name", "Campaign Name (Informational only)"
        )
        ad_group_id_col = find_column(header_map, "Ad Group ID")
        ad_group_name_col = find_column(
            header_map, "Ad Group Name", "Ad Group Name (Informational only)"
        )

        if not all([entity_col, campaign_id_col, campaign_name_col]):
            print(f"  Warning: Missing required columns in '{sheet_name}', skipping")
            continue

        # First pass: collect campaigns
        for row in ws.iter_rows(min_row=2, values_only=False):
            entity = row[entity_col - 1].value
            if not entity:
                continue
            entity_str = str(entity).strip()

            if entity_str == "Campaign":
                cid = row[campaign_id_col - 1].value
                cname = row[campaign_name_col - 1].value
                if cid and cname:
                    cname_str = str(cname).strip()
                    cid_str = str(cid).strip()
                    if cname_str not in campaigns:
                        campaigns[cname_str] = {
                            "campaign_id": cid_str,
                            "sheet_name": sheet_name,
                            "ad_groups": {},
                        }

        # Build reverse lookup for ad group matching
        id_to_name = {
            data["campaign_id"]: cname for cname, data in campaigns.items()
            if data["sheet_name"] == sheet_name
        }

        # Second pass: collect ad groups
        if ad_group_id_col and ad_group_name_col:
            for row in ws.iter_rows(min_row=2, values_only=False):
                entity = row[entity_col - 1].value
                if not entity or str(entity).strip() != "Ad Group":
                    continue

                cid = row[campaign_id_col - 1].value
                ag_id = row[ad_group_id_col - 1].value
                ag_name = row[ad_group_name_col - 1].value
                if cid and ag_id and ag_name:
                    cid_str = str(cid).strip()
                    if cid_str in id_to_name:
                        cname = id_to_name[cid_str]
                        campaigns[cname]["ad_groups"][str(ag_name).strip()] = str(ag_id).strip()

    return campaigns


def read_existing_negatives(wb, sheet_names=None):
    """
    Read existing negative keyword entries from the bulk file for deduplication.

    Returns:
        campaign_negatives: set of (campaign_id, keyword_text_lower, match_type)
        adgroup_negatives: set of (ad_group_id, keyword_text_lower, match_type)
    """
    if sheet_names is None:
        sheet_names = NEGATIVE_KEYWORD_SHEETS

    campaign_negatives = set()
    adgroup_negatives = set()

    for sheet_name in sheet_names:
        if sheet_name not in wb.sheetnames:
            continue

        ws = wb[sheet_name]
        header_map = find_columns_by_header(ws)

        entity_col = find_column(header_map, "Entity")
        campaign_id_col = find_column(header_map, "Campaign ID")
        ad_group_id_col = find_column(header_map, "Ad Group ID")
        keyword_col = find_column(header_map, "Keyword Text", "Keyword or Product Targeting")
        match_type_col = find_column(header_map, "Match Type")

        if not all([entity_col, keyword_col, match_type_col]):
            continue

        for row in ws.iter_rows(min_row=2, values_only=False):
            entity = row[entity_col - 1].value
            if not entity:
                continue
            entity_str = str(entity).strip()

            keyword = row[keyword_col - 1].value
            match_type = row[match_type_col - 1].value
            if not keyword or not match_type:
                continue

            keyword_lower = str(keyword).strip().lower()
            match_type_str = str(match_type).strip()

            if entity_str == "Campaign negative keyword":
                cid = row[campaign_id_col - 1].value
                if cid:
                    campaign_negatives.add((str(cid).strip(), keyword_lower, match_type_str))

            elif entity_str == "Negative keyword":
                ag_id = row[ad_group_id_col - 1].value if ad_group_id_col else None
                if ag_id:
                    adgroup_negatives.add((str(ag_id).strip(), keyword_lower, match_type_str))

    return campaign_negatives, adgroup_negatives


def auto_select_test_campaign(new_rows):
    """Pick the campaign with the fewest new rows for safe testing.

    Args:
        new_rows: list of dicts, each must have 'campaign_id' key

    Returns:
        campaign_id string, or None if no rows
    """
    if not new_rows:
        return None

    campaign_counts = {}
    for r in new_rows:
        cid = r.get("campaign_id")
        if cid:
            campaign_counts[cid] = campaign_counts.get(cid, 0) + 1

    if not campaign_counts:
        return None

    return min(campaign_counts, key=campaign_counts.get)
