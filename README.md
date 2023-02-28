# NAWS Spreadsheet updater

The NAWS import spreadsheet from NAWS no longer imports directly into
BMLT due to changes in the NAWS database as of August 2022.  This
script fixes up some problems so that it can be imported.

Later, we could change the NAWS import code in the BMLT root server to accommodate the new format.



1. Install the PHP spreadsheet package using this command:

`composer install`

2. Run this command:

`php update.php origfile.xlsx newfile.xlsx`

where `origfile.xlsx` is the original spreadsheet file from NAWS.  This will create or overwrite `newfile.xlsx`.

