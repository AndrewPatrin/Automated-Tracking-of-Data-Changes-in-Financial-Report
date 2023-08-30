# Automated Tracking of Data Changes in Financial Report

## Description

This project aims to automate the routine tasks of a SkillFactory finance department employee by processing data from a Google Spreadsheet and generating **XLSX** file with report in the form of two tables. The script keeps track of changes in the original data sheet and update the reports accordingly.

## Requirements

- Python 3.11.x
- [Google Sheets API credentials](https://developers.google.com/workspace/guides/configure-oauth-consent)
- Google Spreadsheet with specific data

## Usage
   1. Clone this repository to your local machine.
   2. Install the required dependencies using `pip`:
        ```bash
        pip install -r requirements.txt
        ```
   3. [Create your Google Cloud Project and follow the instructions](https://developers.google.com/workspace/guides/configure-oauth-consent)
   4. Obtain the necessary Google Sheets API credentials and save them in the project directory `\credentials\credentials.json`.
           *Alternatively, you can utilize the provided credentials within the project.*
   5. Modify the configuration settings in `core\settings.py` to match your Google Sheet and report preferences.
   6. Run the script using the following command:
       ```bash
       python main.py
       ```

## Input Spreadsheet Requirements

The input spreadsheet must include the following columns:

1. "ФИО/Название подрядчика" (Full Name/Contractor Name)
2. "Уникальный номер размещения" (Unique Placement Number)
3. "Дата учета оказания услуг" (Service Provision Accounting Date)
4. "Месяц учета оказания услуг" (Service Provision Accounting Month)

[Example spreadsheet](https://docs.google.com/spreadsheets/d/1crXiHhDA_69vW3cOPMCHETF_l7VDKOMgdaAK9aRoILw/edit#gid=1715990109). While the column order is not crucial, it's essential to maintain the provided column names. If you wish to modify these names, you'll need to update the column names in the `settings.py` file to ensure the script functions correctly.


## Output

The script generates an XLSX format report consisting of two tables. If there are any changes to the service provisioning accounting month or date, the corresponding table will record the modifications.

### A few words from me

While objectively this implementation may not be the most optimal in terms of performance and efficiency, it successfully accomplishes all required tasks. It's important to recognize that this is a first iteration of development and can be improved in the future. This code can serve as a starting point for further optimization and enhancement of the algorithm.

## Contact

If you have any questions or feedback, feel free to contact me at [Telegram](https://t.me/budzzem).
