Attribute VB_Name = "trainpieces"
Public Sub pieces()


Dim currentuser As String
Dim bk As Workbook
'sets the variable to the user that is logged into the computer
currentuser = Environ("username")

' Manage errors & CATIA messages
On Error Resume Next
CATIA.DisplayFileAlerts = False
CATIA.RefreshDisplay = False

   Set CATIA = GetObject(, "CATIA.Application")


'Establish variables

Dim objPart As Variant
Set objPart = CATIA.ActiveDocument.Product

Dim objSel As Selection
Set objSel = CATIA.ActiveDocument.Selection

Dim objPartCollection As Products
Set objPartCollection = objPart.Products

Dim objSubPartCollection As Products

CATIA.RefreshDisplay = False

'Create the Excel report
Set objEXCELapp = CreateObject("EXCEL.Application")
Set bk = Application.ActiveWorkbook
Set sh = bk.Sheets("Track Pieces")

sh.Cells.Select
Selection.ClearContents

'""""""""""""""""""""""""""""""""""""""""""

    'Check number of parts with a type parameter
    objSel.Search ("Name=Type*,all")
    numero = objSel.Count
    
'finds the last cell based on the first row (can change the row by changing "A1" and Cells(1,1) in lines

    If IsEmpty(sh.Range("A1").Value) = False Then
        'Checks to see if the first value is empty, otherwise you return something all the way to the right
        'if it isn't empty, then it runs this
        firstempty = sh.Cells(1, 1).End(xlToRight).Column + 1
    Else
        'if it is empty, it sets the value to 1
        firstempty = 1
    End If
    
    'Assigns the columns that will contain the part information
    labelcol = firstempty
    typecol = labelcol + 1
    
    
    'Sets the header information for each column
    sh.Cells(1, labelcol) = "Part Name"
    sh.Cells(1, typecol) = "Type"
       
   
    'returns the number of parts in the product
    sh.Cells(2, typecol) = numero
    
    'cycle through each capable and returns the name and part type
    For I = 1 To numero
        sh.Cells(2 + I, labelcol) = objPart.Products.Item(I).Name
        sh.Cells(2 + I, typecol) = objPart.Products.Item(I).Parameters.Item("Type").ValueAsString
    Next I
    '""""""""""""""""""""""""""""""""""""""""""
'saves the workbook and tells excel that the workbook is saved to prevent error messages
bk.Save
bk.Saved = True

'objEXCELApp.DisplayAlerts = False
objEXCELapp.Workbooks.Close


CATIA.RefreshDisplay = True
CATIA.DisplayFileAlerts = True

End Sub



