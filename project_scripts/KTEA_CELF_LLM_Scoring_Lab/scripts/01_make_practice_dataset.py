import pandas as pd

rows = [
    (1, "The girl is reading a book.", 2, "obvious"),
    (2, "The girl is reading at the library.", 2, "obvious"),
    (3, "Reading.", 0, "obvious"),
    (4, "Book.", 0, "obvious"),
    (5, "The girl reading a book.", 1, "borderline"),
    (6, "Girl is reading.", 1, "borderline"),
    (7, "The book is reading.", 0, "adversarial"),
    (8, "The girl likes reading.", 1, "adversarial"),
]

df = pd.DataFrame(
    rows,
    columns=[
        "response_id",
        "response_text",
        "human_score",
        "score_type"
    ]
)

df.to_csv(
    "data/scored/practice_responses.csv",
    index=False
)

print(df.head())
print(f"\nCreated {len(df)} responses")