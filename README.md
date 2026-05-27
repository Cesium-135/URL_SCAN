# URL_SCAN
This tool simulates a Chrome browser with human-like click actions to automatically trace and record the full redirection/jump chain of input URLs.  
It is designed for **EN/FR** web pages, supports multi-threading, breakpoint resume, cookie consent handling, and exports both summary and detailed results to Excel with bilingual headers.

## Environment Installation
Before running the script, you need to run the following commands in a bash shell:
```bash
pip install playwright openpyxl matplotlib
```

```bash
playwright install chromium
```

## Key Features
- **Multithreading** – Significantly speeds up scanning (configurable concurrency).  
- **Breakpoint Resume** – Skips already-processed URLs if the output Excel already exists.  
- **Automatic & Click-Triggered Jump Capture** – Records 3xx redirects, client-side navigations, and jumps triggered by clicking on elements (e.g., Espace/Login buttons).  
- **EN/FR Bilingual Support** – All logs, console outputs, Excel headers, and element matching are bilingual.  
- **Cookie Consent Handling** – Automatically accepts cookie popups (main frame + iframes).  
- **Detailed Error Classification** – Extracts error codes (`ERR_*`, HTTP status) and categorises errors into subtypes.  
- **URL Classification** – Tags each final URL as `Public front`, `Espace`, `Authentication`, `Backend/API`, or `Error`, with finer sub-categories.  
- **Statistics & Charts** – Prints distribution and generates a professional chart (BNP Paribas brand colours) with count labels.

## Overall Design
1. Read URL list from an Excel sheet (configurable column, default column A).  
2. Standardise and validate each URL (auto‑add `https://`).  
3. Launch a Chromium browser (headless by default, supports proxy).  
4. For each URL:  
   - Navigate to the page, handle cookie consent popups.  
   - Automatically record all 3xx redirects and client‑side navigations.  
   - If enabled, search for clickable elements (using keyword‑based selectors) and simulate a real user click to uncover hidden jumps (e.g., Espace or login pages).  
   - Capture jumps that open in new windows/tabs.  
5. Merge the complete jump chain for each URL.  
6. Classify the final landing page using rule‑based matching (EN/FR keywords).  
7. Write results to an Excel file with two sheets: **summary** and **jump_chain**.  
8. Generate a statistics chart (main category pie + error code bar) with BNP Paribas brand colours and numeric labels.

## Project Architecture
```
PyCharmMiscProject
├── Results                         # Directory of results
├── matplotlibtest.py               # test of matplot tool
├── README.md                       # this README file
├── prompt.txt                      # the prompt that I use to generate scripts
├── script.py                       # main python script
├── url_list.xlsx                   # input Excel file (e.g., Feuil1 sheet, URLs in column A)
├── URL_SCAN_RESULT_*.xlsx          # output Excel (summary + jump_chain sheets)
├── url_scan_chart_*.png            # generated statistics chart
└── url_scan_YYYYMMDD_HHMMSS.log    # detailed log file
```

## Usage
Run the script from the command line (recommended) or interactively:
```bash
python script.py -i input.xlsx -o output.xlsx
```
If you omit `-i` or `-o`, the script will prompt you to enter the paths interactively.

### Example
```bash
python script.py -i ./url_list.xlsx -o ./result.xlsx
```

## Configuration
All settings are located in the `CONFIG` dictionary at the top of `script.py`.  
You can easily adjust:
- `headless_mode` – set to `False` to see the browser GUI for debugging.  
- `max_workers` – number of concurrent threads (default 30).  
- `timeout` – page timeout in milliseconds.  
- `click_config` – enable/disable click simulation, add more keywords or selectors.  
- `cookie_config` – enable/disable automatic cookie consent.  
- `classify_rules` – add or modify EN/FR keywords for URL classification.  
- `url_column` – change the Excel column that contains URLs (e.g., `"B"`).

## Output
### Summary Sheet (`summary`)
Contains one row per input URL with fields:  
No., Original URL, Standardized URL, Final URL, Jump Count, Full Path, Main Category, Sub Category, Click Triggered, Click Element Info, Auth Required, Status, Error Message, Error Code, Scan Time.

### Detail Sheet (`jump_chain`)
Contains one row per jump step with fields:  
Original URL, Step, URL, From URL, Jump Type, HTTP Status, Trigger Mode, Category, Timestamp.

### Statistics Chart
A PNG file named `url_scan_chart_YYYYMMDD_HHMMSS.png` containing:  
- Pie chart: distribution of main categories (with count & percentage labels).  
- Bar chart (if any errors occurred): distribution of raw error codes (with count labels on top).  
Colours follow **BNP Paribas official brand guidelines** (green primary palette).

## Future Improvements
- Add support for more complex authentication flows (e.g., OAuth redirections).  
- Implement intelligent waiting for single-page applications (SPA) routing.  
- Provide a configuration file (JSON/YAML) instead of hard‑coded settings.  
- Extend classification coverage to more languages (e.g., German, Spanish) upon request.

## Further Improvements
1. ~~The LLM is asked by prompt in Chinese, so it generate code comments and result file in Chinese.
And I asked it to add some English comments, but still remains some place to be changed.~~ SOLVED
2. ~~It cannot handle the situation when a cookie request window is present,
so it need to be improved by adding functions to handle this situation~~ SOLVED
3. It still needs to be compared with manual results to ensure accuracy. TODO
4. Provide a configuration file (JSON/YAML) instead of hard‑coded settings. NEW
5. Extend classification coverage to more languages (e.g., German, Spanish) upon request. NEW
6. As this script developed by LLMs, it's becoming more and more complex. So it needs to be managed as a project. NEW
Otherwise, it will become extremely hard to modify and out of control, especially for users have no knowledge of python.

## License
This tool is a Python demo generated with the assistance of LLMs. Free to use and modify.
