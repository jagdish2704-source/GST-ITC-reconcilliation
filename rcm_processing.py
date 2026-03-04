import pandas as pd


def split_rcm_rows(
    book_df: pd.DataFrame,
    *,
    cgst_nrec_col: str = "CGST_AMT_NREC",
    igst_nrec_col: str = "IGST_AMT_NREC",
    status_col: str = "Status",
    log=None,
):
    """
    Identify RCM rows in Book sheet data and split them out.

    RCM criteria (as requested):
      - CGST_AMT_NREC > 0 OR IGST_AMT_NREC > 0

    Returns:
      (book_without_rcm_df, rcm_df)

    Notes:
      - Matching is case-insensitive + whitespace-tolerant for column names.
      - Sets Status='RCM' on rows that are moved to rcm_df.
    """
    if book_df is None or book_df.empty:
        return book_df, book_df.iloc[0:0].copy()

    df = book_df.copy()

    # Normalize column lookup
    col_lookup = {str(c).strip().lower(): c for c in df.columns if c is not None}
    cgst_col = col_lookup.get(str(cgst_nrec_col).strip().lower())
    igst_col = col_lookup.get(str(igst_nrec_col).strip().lower())

    if not cgst_col and not igst_col:
        # Nothing to do; return an empty RCM df with same headers
        return df, df.iloc[0:0].copy()

    cgst_nrec = pd.Series(0, index=df.index, dtype="float64")
    igst_nrec = pd.Series(0, index=df.index, dtype="float64")

    if cgst_col:
        cgst_nrec = pd.to_numeric(df[cgst_col], errors="coerce").fillna(0)
    if igst_col:
        igst_nrec = pd.to_numeric(df[igst_col], errors="coerce").fillna(0)

    rcm_mask = (cgst_nrec > 0) | (igst_nrec > 0)

    if status_col not in df.columns:
        df[status_col] = ""

    df.loc[rcm_mask, status_col] = "RCM"

    rcm_df = df.loc[rcm_mask].copy()
    non_rcm_df = df.loc[~rcm_mask].copy()

    if log:
        try:
            log(f"RCM rows identified: {int(rcm_mask.sum())}")
        except Exception:
            pass

    return non_rcm_df, rcm_df

