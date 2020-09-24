Attribute VB_Name = "modPicDB"
'---------------------------------------------------------------------------------------
' Module    : modPicDB
' DateTime  : 23/8/2005 16:25
' Author    : Zani
' Purpose   : Show a simpler way to store files on DBs...
'                this sample uses images but it cam be anything
' Note      : The field must be Memo (text) and not OLE (BLOB)
'---------------------------------------------------------------------------------------
Option Explicit
Private chunk() As Byte
'---------------------------------------------------------------------------------------
' Procedure : File2Field
' DateTime  : 23/8/2005 16:31
' Author    : Zani
' Purpose   : Put the File into the field
'---------------------------------------------------------------------------------------
'
Function File2Field(FilePath As String, TField As ADODB.field) As Boolean


   On Error GoTo File2Field_Error

Open FilePath For Binary As 1
ReDim chunk(LOF(1))
Get #1, , chunk()
TField = chunk()
Close 1

File2Field = True
   On Error GoTo 0
   Exit Function

File2Field_Error:
File2Field = False
    MsgBox "File2Field Error " & Err.Number & " (" & Err.Description & ") in procedure File2Field of Módulo modPicDB"
End Function

'---------------------------------------------------------------------------------------
' Procedure : Field2File
' DateTime  : 23/8/2005 16:35
' Author    : Zani
' Purpose   : Get the file out of the field
'---------------------------------------------------------------------------------------
'
Function Field2File(FilePath As String, TField As ADODB.field) As Boolean

   On Error GoTo Field2File_Error

If TField.ActualSize > 0 Then
Open FilePath For Binary As 1
ReDim chunk(TField.ActualSize)
chunk = TField
Put #1, , chunk()
Close 1
End If

Field2File = True
   On Error GoTo 0
   Exit Function

Field2File_Error:
Field2File = False
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Field2File of Módulo modPicDB"
End Function
