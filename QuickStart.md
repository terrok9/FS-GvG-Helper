# QuickStart — FS-GvG-Helper 🛡️⚔️ (Step-by-step)

Follow these steps to generate your **GvG conflict report** as an **Excel (.xlsx)** file.

---

## 1 Prerequisites

### 1.1 Install Python
- Install **Python 3.10+** (recommended).
- Confirm it works:

```bash
python --version
```

### 1.2 Install dependencies
From your project folder:

```bash
pip install requests beautifulsoup4 lxml pandas openpyxl tqdm
```

---

## 2 Prepare your attackers list

### 2.1 Create `attackers.csv`
In the project root, create a file named `attackers.csv` with this format:

```csv
name,level
Alice,5829
Bob,4123
Charlie,1901
Diana,1195
```

✅ Notes:
- `name` is only used for labeling in the report.
- `level` must be an integer.

---

## 3 Authenticate using a Cookie file (recommended)

PowerShell often breaks cookie strings because values may include quotes (`"`) and padding (`==`).
Use `cookie.txt` to avoid escaping issues.

### 3.1 Copy the Cookie header (Chrome/Edge)
1. Log in to **Fallen Sword** in your browser.
2. Press **F12** → open the **Network** tab.
3. Refresh the page (**Ctrl+R**).
4. Click any request that goes to:
   - `https://www.fallensword.com/index.php?...`
5. In the right panel, open **Headers**.
6. Scroll to **Request Headers**.
7. Find **Cookie:** and copy the **value only** (everything after `Cookie:`).

✅ The value usually looks like:

```
fsId=...; fsSessionKey=...; LB="...=="
```

⚠️ Important:
- Copy it from a request to **fallensword.com**, not **account.huntedcow.com**.

### 3.2 Create `cookie.txt`
Create a file named `cookie.txt` in the project root and paste the cookie value as a **single line**:

```txt
fsId=11788; fsSessionKey=XXXXX; LB="YYYY=="
```

---

## 4 Run the scanner

### 4.1 Basic run (A–Z scan + Excel export)
```bash
python gvg_searcher.py --cookie-file cookie.txt --attackers attackers.csv --out conflicts.xlsx
```

When finished, you’ll see:
- `Saved: conflicts.xlsx`
- `Guilds scanned: N`

---

## 5 Understand the Excel output

### 5.1 Sheet: `Conflicts` (main summary)
One row per target guild with:

- `guild_id`, `guild_name`, `guild_url`
- `members_total`
- `active_members_<7d` (active if last activity days `< 7` by default)
- `active_50plus_<7d` (active AND level ≥ 50)
- `min_lvl_active`, `max_lvl_active`
- `viable_attackers` (attackers who have ≥1 valid target in that guild)
- `cap_attacks` (= `viable_attackers * 25`)
- `can_50`, `can_75`, `can_100`
- `recommended` (largest feasible size among 50/75/100)

### 5.2 Sheet: `AttackerCoverage`
Shows for each (guild, attacker):
- attacker level
- how many eligible active targets exist in that guild

### 5.3 Optional Sheet: `Targets` (can be huge)
Lists each eligible active target and which attackers can hit them.

Enable it:

```bash
python gvg_searcher.py --cookie-file cookie.txt --attackers attackers.csv --include-targets --out conflicts.xlsx
```

⚠️ Warning:
- This can generate very large files and run slower.

---

## 6 Example commands (common use cases)

### 6.1 Scan only some letters (faster testing)
```bash
python gvg_searcher.py --cookie-file cookie.txt --attackers attackers.csv --letters A-D --out conflicts.xlsx
```

### 6.2 Stop after N guilds (debug / fast run)
```bash
python gvg_searcher.py --cookie-file cookie.txt --attackers attackers.csv --max-guilds 50 --out conflicts.xlsx
```

### 6.3 Change “active” window (default is <7 days)
```bash
python gvg_searcher.py --cookie-file cookie.txt --attackers attackers.csv --active-days 10 --out conflicts.xlsx
```

### 6.4 Change minimum participants required to initiate a conflict
(Default is 3 = “more than two attackers”)

```bash
python gvg_searcher.py --cookie-file cookie.txt --attackers attackers.csv --min-participants 4 --out conflicts.xlsx
```

### 6.5 Progress bars (v1.2+)
By default, v1.2 shows:
- Outer progress: letters A–Z
- Inner progress: guilds per letter + current guild name

Disable them:

```bash
python gvg_searcher.py --cookie-file cookie.txt --attackers attackers.csv --no-progress --out conflicts.xlsx
```

### 6.6 Slow down requests (safer / more polite)
```bash
python gvg_searcher.py --cookie-file cookie.txt --attackers attackers.csv --min-delay 0.8 --out conflicts.xlsx
```

---

## 7 Troubleshooting

### 7.1 “Not logged in”
- Cookie may be expired or incomplete.
- Ensure the cookie was copied from **fallensword.com** requests (not huntedcow account domain).
- Re-copy cookie and replace `cookie.txt`.

### 7.2 PowerShell: `unrecognized arguments ...`
This happens when you pass cookies directly as CLI text and PowerShell breaks quotes.
✅ Use `--cookie-file cookie.txt`.

### 7.3 `Targets` is too slow / too big
- Run without `--include-targets`.
- Limit scan:
  - `--letters A-B`
  - `--max-guilds 25`

### 7.4 Some members missing activity days
If `Last Activity` is missing from the tooltip, `inactive_days` is `None` and the tool may ignore them for “active” filtering.

---

## 8 Security & Responsible Use (very important)

### 8.1 Anti-ban / rate-limit hygiene
- Keep a non-zero delay (`--min-delay`).
- Do not run multiple instances at once.
- Avoid scanning A–Z repeatedly in short time windows.
- If errors increase, stop and increase delay.

### 8.2 Cookie safety
Treat `cookie.txt` like a password:
- Do not share it.
- Do not commit it to git.
- Add to `.gitignore`:

```txt
cookie.txt
*.cookie.txt
```

If you suspect compromise:
- regenerate the cookie by logging in again
- consider changing your password

### 8.3 Respect Terms of Service
You are responsible for using this tool in compliance with Fallen Sword’s Terms of Service and community rules.
