# Coaching Guide — Pan-India Site Generator

## What's in this folder
- `generate_site.py` — the generator script
- `style.css` — shared stylesheet (needed alongside the script)
- `institutes_master.xlsx` — your data (download separately, keep alongside this script)
- `output_sample/` — a sample of what running the script produces (already generated from your real Lucknow data, for reference)

## One-time setup
```powershell
pip install openpyxl --break-system-packages
```
(On Windows without the flag: `pip install openpyxl`)

## How to (re)build the whole site
1. Add/edit institute rows in `institutes_master.xlsx` (see the "How to add a city" tab inside it)
2. Save the spreadsheet
3. Run:
   ```powershell
   python generate_site.py institutes_master.xlsx output
   ```
4. The complete site — every state, every city, every exam page — is generated fresh into the `output/` folder
5. Drag `output/` onto vercel.com/drop (or push it to GitHub) to publish

## Adding a new city or state
No code changes needed — just add rows to the spreadsheet:
- Type the **State** and **City** exactly as you want them displayed (e.g. `Bihar`, `Patna`)
- Fill in **Exams** (comma list: `SSC,BankPO,UPSC`) to control which exam pages the institute appears on
- Leave fee/coverage cells blank for exams the institute doesn't offer
- Re-run the script — the new city's folder and pages are created automatically, and it's added to the homepage city picker

## Featuring an institute (rent model)
Type the exam name (e.g. `SSC`) into that institute's **"Featured On"** column, save, and re-run the script.
That institute gets a highlighted row + badge and becomes the default lead-routing target on that exam page.
You can also preview before committing by adding `?featured=institute-slug` to any exam page's URL.

## URL structure
```
/index.html                                   → state/city picker
/<state-slug>/<city-slug>/index.html          → city homepage
/<state-slug>/<city-slug>/ssc-coaching-<city-slug>.html
/<state-slug>/<city-slug>/bank-po-coaching-<city-slug>.html
/<state-slug>/<city-slug>/upsc-coaching-<city-slug>.html
```
This matches how the site scales to any number of states/cities without manual folder work.
