# Dashlane Flagged Password Mail Merge Puller
A PowerShell script that reads a CSV export of user password health scores the Dashlane Admin > Dark Web Insights > Employee emails CSV file, filters out users with poor scores, and generates a clean Excel sheet (`dashlaneflags_mailmerge.xlsx`) for use in a mail merge email.

---

## Features

- Reads from `team-members.csv`
- Sorts users by `password_health_score`
- Filters users with numeric scores **below 60**
- Ignores empty or invalid entries
- Extracts email addresses from column `login email`
- Generates a new Excel file with:
  - `FirstName` (derived from email prefix)
  - `Email`
- Automatically creates the output file in the same directory as the script

---

## Input Format

Your `team-members.csv` file must contain at least the following columns:

| Column Name         | Example Value              | Description |
|----------------------|----------------------------|--------------|
| `login email`        | `jane.doe@company.com`     | User's login email address |
| `password_health_score` | `45`                    | Numeric password strength score (0–100) |

Example CSV:

```csv
login email,password_health_score
jane.doe@company.com,45
alex.smith@company.com,78
pat.taylor@company.com,55
