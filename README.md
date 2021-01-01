# CSV-Parser-Function-Tunggal__VB6
Bahan Ajar Fundamental Pemrograman Visual Basic 6.0 - CSV Parser Function Tunggal<br><br>

'------------------------------------------------------------------------------
'   CSV Parser
'   This class handles retrieving elements from a CSV (C_omma S_eperated V_alues)
'   file. In the CSV file each line is a record and each field in the record is
'   seperated from its neighbor by a delimiter character. The character is usually
'   a comma (,) but can be any character.
'
'   This class requires a reference to the MS Scripting Runtime.
'
'   Create an instance of the class (Dim CSVP as New CSVParse)
'   Set the FieldSeperator property if it is not comma.
'   Set the FileName property using the full path to the target file.
'      a. Read the Status property. If it is false, the file was not
'         accessed so call the GetErrorMessage function to retrieve the
'         descripition of the problem
'   Process the file as follows:
'
'       While CSVP.LoadNextLine = True
'           MyString = CSVP.GetField(n) <- for each field you want to read
'                                          where n is the field number where
'           .                              1 is the first field.
'           .
'           .
'       Wend
'----------------------------------------------------------------------------<br><br>

Lihat Source Code : <br>
- <a href="https://github.com/RizkyKhapidsyah/CSV-Parser-Function-Tunggal__VB6/blob/main/clsCSV.cls">Program</a>
