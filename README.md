# Combine-ER-Reports

Purpose: Combine weekend ER reports to save time and potential manual copy/paste mistakes.

## Why Important

After weekends, we work the reports that were run on the weekend as well as current day. Therefore we combine (append) all files and work mutiple days as 1 file.

Combining files is a time consuming menial task with lots of room for manual copy/paste errors and something my work team needs to do on a regular basis (usually every Monday).

Time saved per Monday not manually combining 30+ files is around 20-25 minutes.

## Usage

Put all raw ER report files (.xls) into directory User/Desktop/python_projects/combine_er_reports/data/raw

Combined files (.xlsx) can be saved to directory User/Desktop/python_projects/combine_er_reports/data/exports

## Installation

Use the package manager [pip](https://pip.pypa.io/en/stable/) to install respective libraries.

```bash
pip install pandas
pip install numpy
pip install glob
pip install openpyxl
```

## Contributing

I was helped immensely by https://github.com/Fehiroh who walked me through the logic and thought process of the script (originally done in R/tidyverse). I then re-worked the script into Python/Pandas using the same logic. 
