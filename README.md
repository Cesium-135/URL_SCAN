# URL_SCAN
This is a python demo of automatically tracing and recording the jump chains of input URLs,
by simulating a chrome browser and human-like click.
With the help of and generate by Chinese LLMs.

## Environment Installation
Before running the script, you need to run the following commands in a bash shell:
```bash
pip install playwright openpyxl
```

```bash
playwright install chromium
```

## Overall Design
1. Read in the URL list in an Excel sheet
2. Standardize and pre-process URL formats
3. Launch the simulated chrome browser and create tabs with input URL
4. Access URL and capture automatic or manual jump
   - a. For automatic jumps, simply record it
   - b. For non-automatic jump website, positining elements by keywords, and simulate click it
   - c. Capture full jump chain automatic or triggered by click
5. Merge complete jump chain for each input URL
6. Classification and tagging results according to rules
7. Write results back to final Excel file.

## Project Architecture
```
PyCharmMiscProject
├── README.md               # this file
├── prompt.txt              # the prompt that I use to generate scripts
├── script.py               # main python script
├── url_list.xlsx           # input URL list file
└── url_jump_result.xlsx    # output file
```

## Further Improvements
1. The LLM is asked by prompt in Chinese, so it generate code comments and result file in Chinese.
And I asked it to add some English comments, but still remains some place to be changed.
2. It cannot handle the situation when a cookie request window is present,
so it need to be improved by adding functions to handle this situation
3. It still needs to be compared with manual results to ensure accuracy.
