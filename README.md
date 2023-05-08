NOTE: For any of these scripts to function properly, xlwings and PyMuPDF must be installed. 


The BOM transformation and ID replacement folder is for scripts that use xlwings to manipulate BOMs (Excel files).

    All files used in these scripts should be stored in the "templates" folder, or else their paths must be changed.

    1. BOM transformations.py 
    
        Goal: To transform the formatting of the existing BOMs to fit a specific template based on the 32 steps listed in templates/BOM template.xlsx (second sheet).
    
    2. HDG BOM description replacement.py

        Goal: To standardize HDG metric bolts by replacing all HDG metric bolt IDs and descriptions with a specified format. 

        2 standards exist for this: ASTM and ISO. ISO fasteners have screws (fully threaded) and bolts (partially threaded).
        All IDs should be replaced with the corresponding ASTM/ISO ids. If a fully threaded bolt is listed, a note must be added to the BOM indicating that specific ID is threaded.
    
    3. replacing descriptions in BOM.py

        Goal: To replace, delete, or update the IDs or descriptions for a specific series of items.

        The data used for these replacements, including old ID, new ID, Eng Dash description, and DWG description are all stored in templates/S200 unique ALL.xlsx. Of course, a different file can be used to store this, as long as the paths are updated accordingly.


The pdf replacement folder contains 2 files:

    1. dictionary.py, which maps old HDG IDs to new HDG IDs
    
        Goal: To create a dictionary mapping the old HDG IDs (used in BOM before HDG ID updates) to new IDs (used in new BOM after HDG ID updates).

        This is important because the keys of the dictionary (old IDs) are used to search in the PDF for IDs to replace, while the values store the new, correct IDs to be used in the drawings.
        Also collects a list of sheet names in the BOM, as they correspond to the drawing names and enables dynamic loading and editing of all the PDFs in a series.

    2. HDG pdf replacement.py

        Goal: To replace all old IDs redlined in Adobe Acrobat comments in each of the drawings in a series with new IDs, both of which are stored in the "old_to_new" dictionary imported from dictionary.py

        Please note the paths defined in the script for the PDF location and new location for the PDF to be saved. 
