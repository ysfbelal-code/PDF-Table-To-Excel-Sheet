def _merge_multiline_generic(df):
    """
    Merge multi‑line content that was split across rows.
    1) Consecutive rows with exactly one non‑null cell (same column, text) are merged into one row.
    2) Remaining single‑cell rows are merged into the nearest preceding anchor row (>=2 non‑null).
    """
    if df.empty or len(df) < 2:
        return df

    # Helper: check if value is data‑like (date, amount, ID)
    def _is_data_like(val):
        if val is None or not isinstance(val, str):
            return False
        s = val.strip()
        if re.search(r"\d{2}[-/]\d{2}[-/]\d{2,4}", s):      # date
            return True
        if re.search(r"\d+[\.,]\d{2}", s):                  # amount
            return True
        if re.match(r"^[\d\-\s]+$", s) and re.search(r"\d", s):  # numeric ID
            return True
        return False

    # ---- Stage 1: merge consecutive single‑cell rows ----
    # Identify rows that are "single‑cell" (only one non‑null and not data‑like)
    single_indices = []
    for idx in range(len(df)):
        row = df.iloc[idx]
        non_null = row.notna()
        if non_null.sum() == 1:
            col = row[non_null].index[0]
            val = row[col]
            if not _is_data_like(val):
                single_indices.append((idx, col))

    # Group consecutive indices with the same column
    groups = []
    if single_indices:
        groups = []
        current_group = [single_indices[0]]
        for i in range(1, len(single_indices)):
            prev_idx, prev_col = single_indices[i-1]
            curr_idx, curr_col = single_indices[i]
            # Consecutive and same column
            if curr_idx == prev_idx + 1 and curr_col == prev_col:
                current_group.append(single_indices[i])
            else:
                groups.append(current_group)
                current_group = [single_indices[i]]
        groups.append(current_group)

    # For each group, merge the text and replace the group with a single row
    rows_to_keep = []
    skip_indices = set()
    for group in groups:
        if len(group) >= 2:
            # Merge values
            merged_val = "\n".join([str(df.iloc[idx][col]) for idx, col in group if pd.notna(df.iloc[idx][col])])
            # Keep the first index of the group to hold merged value, mark others to drop
            first_idx, first_col = group[0]
            # Update the first row with merged value
            df.iloc[first_idx, first_col] = merged_val
            # Mark the rest to be dropped
            for idx, col in group[1:]:
                skip_indices.add(idx)
        else:
            # Single row group: keep it (no merge needed)
            pass

    # Drop rows marked for skipping
    if skip_indices:
        df = df.drop(index=list(skip_indices)).reset_index(drop=True)

    # ---- Stage 2: merge remaining single‑cell rows into nearest anchor ----
    # Re‑compute non‑null counts after stage 1
    non_null_counts = df.notna().sum(axis=1)
    candidate_indices = []
    for idx in non_null_counts[non_null_counts == 1].index:
        row = df.loc[idx]
        col = row.first_valid_index()
        if col is not None and not _is_data_like(row[col]):
            candidate_indices.append((idx, col))

    if not candidate_indices:
        return df

    anchor_indices = non_null_counts[non_null_counts >= 2].index.tolist()
    if not anchor_indices:
        return df

    # Merge candidates into nearest preceding anchor
    merged_rows = {}
    current_anchor = None
    for idx in range(len(df)):
        if idx in anchor_indices:
            current_anchor = idx
            merged_rows[idx] = df.iloc[idx].copy()
        elif idx in [c[0] for c in candidate_indices] and current_anchor is not None:
            col = [c for c in candidate_indices if c[0] == idx][0][1]
            if col is not None:
                anchor_row = merged_rows[current_anchor]
                if pd.notna(anchor_row[col]):
                    anchor_row[col] = str(anchor_row[col]) + "\n" + str(df.iloc[idx][col])
                else:
                    anchor_row[col] = str(df.iloc[idx][col])

    if merged_rows:
        sorted_indices = sorted(merged_rows.keys())
        new_df = pd.DataFrame([merged_rows[i] for i in sorted_indices], columns=df.columns)
        new_df.dropna(how='all', inplace=True)
        return new_df
    else:
        return df
