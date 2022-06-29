# COSHH_form
Save all files to the same directory.
Edit the COSHHdata.xlsx file to have the details of the chemical procedure.  
Run the COSHH_fill2.py script
A <proceeduretitle>.docx file will be produced in the folder.
If the substances have been found in the wikipedia (with hazards codes) then the harzards will be included in the document.
The other files e.g. COSHH_template do not need to be opened, they are used to generate the end document.
The script uses the doc_merge module and beautifulsoup along with other packages.
