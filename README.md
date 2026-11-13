# Big-Little Matching

This script reads the ranking matrix CSV files and generates up to 10 pairing suggestions in an Excel workbook, with one worksheet per suggestion.

## Run

```bash
python biglittle.py "little preferences.csv" "big preferences.csv"
```

The script writes `suggested pairings.xlsx`. Each suggestion assigns every little to exactly one big. Two different bigs receive two littles each; all other bigs receive one. By default, the two most-requested bigs are selected to take twins.

To specify twin-taking bigs, edit `TWIN_BIG_NAMES` near the top of `biglittle.py`, for example:

```python
TWIN_BIG_NAMES = ["Random", "Guy"]
```

Enter one or two names from the big preference CSV. If you enter one name, the script pairs that big with the most-requested other big. Leave the list empty (`[]`) for automatic selection of both.

All ranked options are used, including both names in `(A)`/`(B)` cells. Suggestions minimize the combined ranks from both sides. Unranked pairings count below ranked pairings. Each worksheet includes preference ranks and whether both people ranked each other. The `Twin Big` column is populated only on rows belonging to a twin-taking big.
