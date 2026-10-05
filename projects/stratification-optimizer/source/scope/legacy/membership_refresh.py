"""Refresh saved membership from current eligible scored records."""
import pandas as pd


def refresh_membership(previous, universe, clean_id):
    previous = previous.copy()
    universe = universe.copy()
    if "teid" not in previous or "teid" not in universe:
        raise ValueError("Sample refresh requires a TEID column in both sources.")
    for frame in (previous, universe):
        frame["teid"] = frame.teid.map(clean_id)
        if frame.teid.isna().any() or frame.teid.duplicated().any():
            raise ValueError("Sample refresh requires nonmissing, unique TEIDs; resolve duplicates before running.")
    selected_ids = set(previous.teid)
    current = universe.set_index("teid", drop=False)
    audit = []
    fields = [col for col in previous.columns if col in universe and col != "teid"]
    for _, old in previous.iterrows():
        teid = old.teid
        if teid not in current.index:
            audit.append({"teid": teid, "status": "missing_or_ineligible", "field": "", "previous": "", "updated": ""})
            continue
        new = current.loc[teid]
        changes = []
        for col in fields:
            a, b = old[col], new[col]
            if (pd.isna(a) and pd.isna(b)) or str(a).strip() == str(b).strip():
                continue
            changes.append({"teid": teid, "status": "changed", "field": col, "previous": a, "updated": b})
        audit.extend(changes or [{"teid": teid, "status": "unchanged", "field": "", "previous": "", "updated": ""}])
    master = universe.loc[universe.teid.isin(selected_ids)].copy()
    available = universe.loc[~universe.teid.isin(selected_ids)].copy()
    return master, available, pd.DataFrame(audit, columns=["teid", "status", "field", "previous", "updated"])
