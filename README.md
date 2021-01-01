# CSV-Parser-Function-Tunggal__VB6
Bahan Ajar Fundamental Pemrograman Visual Basic 6.0 - CSV Parser Function Tunggal<br><br>

'---<br>
'   CSV Parser<br>
'   This class handles retrieving elements from a CSV (C_omma S_eperated V_alues)<br>
'   file. In the CSV file each line is a record and each field in the record is<br>
'   seperated from its neighbor by a delimiter character. The character is usually<br>
'   a comma (,) but can be any character.<br><br>
'
'   This class requires a reference to the MS Scripting Runtime.<br><br>
'
'   Create an instance of the class (Dim CSVP as New CSVParse)<br>
'   Set the FieldSeperator property if it is not comma.<br>
'   Set the FileName property using the full path to the target file.<br>
'      a. Read the Status property. If it is false, the file was not<br>
'         accessed so call the GetErrorMessage function to retrieve the<br>
'         descripition of the problem<br>
'   Process the file as follows:<br><br>
'
'       While CSVP.LoadNextLine = True<br>
'           MyString = CSVP.GetField(n) <- for each field you want to read<br>
'                                          where n is the field number where<br>
'           .                              1 is the first field.<br>
'           .<br>
'           .<br>
'       Wend<br>
'---<br><br>

Lihat Source Code : <br>
- <a href="https://github.com/RizkyKhapidsyah/CSV-Parser-Function-Tunggal__VB6/blob/main/clsCSV.cls">Program</a>
