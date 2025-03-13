This is a simple storage system, with the options to read XML files with products, Create an Excel File based on the database and add/remove products.

The database used is PostgreSQL 17 with PgAdmin 4.
Put the access keys on your config.json, I uploaded an example to help understanding.

For the interface I used TkInter.

use "pyinstaller --onefile --windowed --icon=content\icone-storage.ico Estoque.py" with your venv active to create an executable of the program.

the log.log contains the prints of the app.
The content folder contains the icon of the project.


It will create 3 folders:
xmls:put the xml files you want to read;
xmls_old:When read is done, the application moves all xmls to this folder;
excel: When the button "Gerar Excel" is clicked, this is the dir where the excel file is generated.
