import pandas as pd
import pytest

from scope.legacy.membership_refresh import refresh_membership


def test_refresh_keeps_ids_but_replaces_old_values_and_excludes_missing_cases():
    old = pd.DataFrame({"teid": ["1", "2"], "tmp_EBD": [0, 0], "score": [10, 20]})
    new = pd.DataFrame({"teid": ["1", "3"], "tmp_EBD": [1, 0], "score": [99, 30]})
    master, pool, audit = refresh_membership(old, new, str)
    assert master.teid.tolist() == ["1"]
    assert master.score.tolist() == [99]
    assert master.tmp_EBD.tolist() == [1]
    assert pool.teid.tolist() == ["3"]
    assert audit.loc[audit.teid.eq("2"), "status"].tolist() == ["missing_or_ineligible"]


def test_refresh_rejects_ambiguous_ids():
    old = pd.DataFrame({"teid": ["1"]})
    new = pd.DataFrame({"teid": ["1", "1"]})
    with pytest.raises(ValueError, match="unique TEIDs"):
        refresh_membership(old, new, str)
