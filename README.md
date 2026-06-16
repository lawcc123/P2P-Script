# P2P-Script

`P2P_Registry_Sync.py` is a tailor-made automation script in order to update the excel payment record spreadsheet created by my manager.

-Sorting function is not allowed to use in order to maintain the visual structure.
-At least one empty row between different contractor to make it visually separate.
-The script updates the cheque registry workbook with P2P status, IR number, cheque number, and cheque date to excel worksheet from Yardi P2P by performing HTTP API GET request
-Please note that the record will only be updated once it has invoice and PO number or the search will fail.
-You have to sign in manually in order to avoid bot detection from Cloudfare.
-Once sign in, click "Elevate", the script will start to retrieve the details.


`P2P_PO_Create.py` creates Yardi P2P purchase orders from rows in the cheque registry workbook.

-To create a PO, you must enter Contractor name, invoice number, invoice amount, account code and add "Require to create PO" on P2P status column. PO creation will fail it lacks one of this.
-The script does not approve the PO, add attachement and advance to manager automatically, please complete this step on workflow dashboard.
-Once it is created successfully, the script will automatically fill in the PO number column and update the P2P status.
