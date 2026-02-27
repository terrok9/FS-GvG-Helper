# FS-GvG-Helper 🛡️⚔️  
**Fallen Sword Guild-vs-Guild Conflict Finder + Excel Export**

FS-GvG-Helper is a small Python CLI tool that scans the **A–Z guild directory** in Fallen Sword, opens each guild page, reads **member levels + last activity**, and produces an **Excel spreadsheet** showing which guilds are viable conflict targets for your attacker roster.

It’s built for quick planning: “Given our available attackers, which guilds can we legally conflict and which size (50/75/100) is feasible?”

**Done for people that had burned time looking for guilds**

## 🚀 Start here
Read the step-by-step guide: **[QuickStart.md](QuickStart.md)**

---

## ✨ Features

- ✅ **A–Z guild scanning** (with paging per letter)
- ✅ Parses guild member list:
  - Player name + level
  - “Last Activity” (days) from the `data-tipped` tooltip
- ✅ Filters targets to **active players** (default `< 7 days inactive`)
- ✅ Computes **valid attacks** with your rule:  
  **the lowest player level sets the allowed ± range**
- ✅ Calculates feasibility for conflict sizes **50 / 75 / 100**
  - Contribution cap: **25 attacks per attacker**
  - Minimum participants configurable (default **3**, i.e., “more than two attackers”)
- ✅ Exports a **friendly `.xlsx` report**
  - `Conflicts` summary sheet
  - `AttackerCoverage` per guild coverage sheet
  - Optional `Targets` sheet (can be large)
- ✅ Supports **cookie-file auth** (recommended) to avoid shell-escaping issues
- ✅ Optional CLI **progress bars** (letters + current guild) in v1.2+

---

## 🧠 How it works (high level)

1. **Authenticate** (recommended: browser Cookie header → `cookie.txt`)
2. For each letter **A → Z**:
   - Fetch the A–Z guild list pages
   - Extract `guild_id`, guild name, and member count
3. For each guild:
   - Open the guild page and parse the members table
   - Extract each member’s:
     - `level`
     - `Last Activity: Xd ...` from the tooltip (cached by the game)
4. Keep **active targets** only: `inactive_days < 7` (configurable)
5. For each attacker and each active target:
   - Check **attack eligibility** using:
     - `low = min(attacker_level, target_level)`
     - `delta = range_delta(low)` per ranges below
     - Eligible if: `abs(attacker_level - target_level) <= delta`
6. Count how many attackers have ≥1 eligible target in that guild.
7. Compute max possible conflict attacks: `viable_attackers * 25`
8. Determine if you can run:
   - **50 attacks** (need enough attackers to cover 50 with 25 each + min participants)
   - **75 attacks**
   - **100 attacks**
9. Export results to Excel.

---

## 📏 Level-range rules implemented

Based on the “Important Notes on Guild Conflicts” table:

- 50–300: ±25  
- 301–700: ±50  
- 701–1000: ±100  
- 1001–2000: ±125  
- 2001–3000: ±150  
- 3001–4000: ±175  
- 4001–5000: ±200  
- etc… (+25 per additional 1000 levels)

**Important**: Range is based on the **lowest** level between attacker and defender.

---

## ✅ Requirements

- Python 3.10+ recommended
- Works on Windows / Linux / macOS

Install dependencies:

```bash
pip install requests beautifulsoup4 lxml pandas openpyxl tqdm

```markdown
## 🔐 Security & Responsible Use (anti-ban + account safety)

This tool behaves like a browser client and makes many requests. Use it carefully to reduce the chance of rate-limits, flags, or account compromise.

### ✅ Don’t get flagged / don’t get banned
- **Respect rate limits**: keep a non-zero delay between requests (`--min-delay`).
- **Do not run multiple instances** at once.
- Avoid scanning **A–Z repeatedly** in a short time window. If you’re iterating, scan a smaller scope:
  - `--letters A-D`
  - `--max-guilds 50`
- If the site slows down or errors increase, **stop** and increase delay.

### ✅ Prefer cookie-file auth (recommended)
Passing cookies directly in CLI can break (PowerShell quoting) and can leak into shell history.
Use:
```bash
python gvg_searcher.py --cookie-file cookie.txt --attackers attackers.csv --out conflicts.xlsx