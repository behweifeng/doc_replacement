# doc_replacement
#### Create multiple word documents with keywords being replaced (words in header will be replaced as well)

1. Create a .docx file as a template, giving each of the placeholders a 
unique name

2. Create a .xlsx file with row 0 being the placeholders that you want to replace, and subsequent rows be your values (each case). Each row will create one document. 

3. Navigate to the folder in terminal / cmd

4. Run the following command (where document.docx is your .docx file name and excel.xlsx is your .xlsx file name)
> python3 test.py document.docx excel.xlsx 

5. The resulting documents will be created as result_.docx 


#### Note: you will need to pre-install the following libraries

pandas
> python3 -m pip install pandas

openpyxl
> python3 -m pip install openpyxl

python-docx
> python3 -m pip install python-docx
