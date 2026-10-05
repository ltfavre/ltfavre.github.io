# BASC legacy regression fixture

This is the small, synthetic fixture previously used to verify the known-working
SOE Part A/Part B implementation. The six Parquet files stand in for SAS files
with the same base names; they are enabled only when the explicit
`allow_parquet_fixtures` test flag is true. Production runs still require the
configured `.sas7bdat` inputs.

Known legacy invariants for this fixture:

- Part A starting/master pool: 2,246 TEIDs
- Part A available pool: 814 TEIDs
- Candidate universe: 3,060 TEIDs
- Final selected sample: 2,400 TEIDs
- Final available pool: 660 TEIDs
- All 11 form-by-age cells meet exact N and exact 50/50 sex balance

The eight-worker CP-SAT search may choose among equivalent swap sets. Therefore
the regression test compares every public SCOPE artifact to the raw legacy
artifact from the same run, while comparing the counts and hard constraints to
the fixed known-working baseline above.
