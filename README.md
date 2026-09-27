# Tableau User & Group Report

A small Python script that connects to a Tableau Cloud (or Tableau Server) site and exports every user and group into a clean, formatted Excel report.

It is useful for access reviews, licence housekeeping, and quickly answering questions like "which groups is this person in?" or "who is in this group?".

## What you get

The script produces one Excel file (for example `output/Tableau_User_Group_Report_20260927_143000.xlsx`) with three sheets.

**Stats** shows the server and site, when the report was generated, headline counts (total users, total groups, empty groups, users with no group, users in two or more groups) and a breakdown of users by site role.

**User Summary** lists one row per user with display name, username, email, site role, all groups they belong to, and last login date. Users who are not in any group are highlighted in yellow.

**Group Details** lists one row per group member with the group name repeated on every row, so the sheet can be filtered or pivoted easily. Empty groups are still listed so nothing is missed.

## Requirements

You need Python 3.9 or newer, a Tableau site administrator account (or an account with permission to view users and groups), and a Personal Access Token (PAT) for that account.

To create a PAT in Tableau, open **My Account Settings**, go to **Personal Access Tokens**, give the token a name and click **Create**. Copy the secret straight away, because Tableau only shows it once.

## Setup

**1. Get the code**

```bash
git clone https://github.com/<your-username>/tableau-user-group-report.git
cd tableau-user-group-report
```

**2. Install the dependencies**

```bash
pip install -r requirements.txt
```

**3. Create your config file**

Make a copy of the example config and name it `config.json`:

```bash
# macOS / Linux
cp config.example.json config.json

# Windows
copy config.example.json config.json
```

Open `config.json` and fill in your details:

| Setting | What to enter |
|---|---|
| `server_url` | Your Tableau URL, for example `https://10ax.online.tableau.com` |
| `site_content_url` | The site name as it appears in the URL after `/site/` |
| `pat_name` | The name of your Personal Access Token |
| `pat_secret` | The secret of your Personal Access Token |
| `api_version` | Tableau REST API version. `3.21` works for current Tableau Cloud; change it if your server needs another version |
| `timezone` | Optional. A timezone name such as `Europe/London` or `America/New_York` for the report timestamp. Leave empty to use your computer's local time |
| `output_dir` | Folder where reports are saved. Defaults to `output` |

`config.json` is listed in `.gitignore`, so your credentials stay on your machine and are never pushed to GitHub.

## Run it

```bash
python tableau_user_group_report.py
```

To use a config file stored somewhere else:

```bash
python tableau_user_group_report.py --config path/to/my_config.json
```

The script prints its progress and tells you where the report was saved.

## Project structure

```
tableau-user-group-report/
├── tableau_user_group_report.py   Main script
├── config.example.json            Template for your connection details
├── config.json                    Your real details (you create this; not committed)
├── requirements.txt               Python dependencies
├── .gitignore                     Keeps secrets and reports out of Git
├── LICENSE                        MIT License
└── README.md
```

## How it works

The script signs in to Tableau with your PAT using the Tableau REST API, fetches all users on the site, fetches all groups, then fetches the members of each group. It handles pagination automatically, so large sites are fully covered. After collecting the data it signs out and builds the Excel file with openpyxl.

It only reads data. Nothing on your Tableau site is created, changed or deleted.

## Notes

Tableau includes a built-in **All Users** group that every user belongs to, so "Users with no group" will usually be zero.

Group membership is fetched one group at a time, so on sites with many groups the script can take a few minutes.

Keep your PAT secret safe. If you think it has been exposed, revoke it in Tableau and create a new one.

## Troubleshooting

**HTTP 401 on sign in:** the PAT name or secret is wrong, has expired, or has been revoked. Also check that `site_content_url` is correct.

**HTTP 404 or version errors:** check `server_url` and try a different `api_version` supported by your Tableau version.

**Config file not found:** make sure `config.json` is in the same folder you run the script from, or pass its location with `--config`.

## License

This project is released under the MIT License. See [LICENSE](LICENSE) for details.
