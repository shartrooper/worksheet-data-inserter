# worksheet-data-inserter
Use both PyPDF2 and openpyxl third party modules. Descriptive header in .py file



**Features:**

- Multiple pdf/txt loading.
- Read and extract data streams from even corrupted Pdfs.
- Directory route file search, avoid unintentional file overwriting.
- Openpyxl data inserting and format.
- If the column threshold is reached, the app switches into another worksheet or create a new worksheet if there's none without empty available columns, and rename Its title with the  reports date range.



**TODO:**

- After loading a workbook from the user input prompt (not from script exec), create a new worksheet if current report's date is well over a month of difference to prior  worksheet's dates with available column space.