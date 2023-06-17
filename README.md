The repo has following files :-
1. output folder for storing zipfiles and ans.csv files created after running code.py
2. src/extractpdf contains code.py
3. TestDataSet folder contains test output.pdf files
4. ans.csv file contains data obtained after extracting 100 test output pdf
5. pdfservices-api-credentials.json and private.key file contains credentials
6. Enter your credentials in pdfservices-api-credentials.json and paste private.key to execute extraction

In code.py, utitlised PDF-EXTRACT-API Python to extract data,

Took advantage of 'Bounds' and 'Text attribute in JSON file present inside zipfile after extraction and converted into CSV with the help of pandas dataframe
