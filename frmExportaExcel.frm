VERSION 5.00
Begin VB.Form frmExportaExcel 
   Caption         =   "Exporta a Excel"
   ClientHeight    =   4980
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6975
   LinkTopic       =   "Form1"
   ScaleHeight     =   4980
   ScaleWidth      =   6975
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   5520
      TabIndex        =   0
      Top             =   360
      Width           =   1215
   End
End
Attribute VB_Name = "frmExportaExcel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ------------------------------------------------------------------------------------
' \\ -- Botón para exportar
' ------------------------------------------------------------------------------------
Private Sub Command1_Click()
  Dim sPathDB As String
  Dim Consulta As String

  ' -- Path de la base de datos
  sPathDB = "C:\"

  ' -- Cadena Sql
  Consulta = "Select Distinct Modelo From Geometricas Order By Modelo"

  ' -- Enviar el Path de la base de datos y la consulta sql
  If Exportar_ADO_Excel(sPathDB, Consulta, "c:\libro.xLS") Then
    MsgBox "Datos Exportados a Excel Correctamente", vbInformation
  End If
End Sub

' ------------------------------------------------------------------------------------
' \\ -- Función para exportar el recordset ADO a una hoja de Excel
' ------------------------------------------------------------------------------------
Private Function Exportar_ADO_Excel(sPathDB As String, Sql As String, sOutputPathXLS As String) As Boolean

  On Error GoTo errSub

  Dim cn As New ADODB.Connection
  Dim rec As New ADODB.Recordset
  Dim Excel As Object
  Dim Libro As Object
  Dim Hoja As Object
  Dim arrData As Variant
  Dim iRec As Long
  Dim iCol As Integer
  Dim iRow As Integer

  Me.Enabled = False

  ' -- Abrir la base
  ' cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & sPathDB & ";"

  cn.Open "Provider=SQLOLEDB; " & _
          "Initial Catalog=SudokuGeneral; " & _
          "Data Source=LAPTOPMIGUEL\SQLEXPRESS; " & _
          "integrated security=SSPI; persist security info=True;"

  ' -- Abrir el Recordset pasándole la cadena sql
  rec.Open Sql, cn

  ' -- Crear los objetos para utilizar el Excel
  Set Excel = CreateObject("Excel.Application")
  Set Libro = Excel.Workbooks.Add

  ' -- Hacer referencia a la hoja
  Set Hoja = Libro.Worksheets(1)

  Excel.Visible = True: Excel.UserControl = True
  iCol = rec.Fields.Count
  For iCol = 1 To rec.Fields.Count
    Hoja.Cells(1, iCol).Value = rec.Fields(iCol - 1).Name
  Next

  If Val(Mid(Excel.Version, 1, InStr(1, Excel.Version, ".") - 1)) > 8 Then
    Hoja.Cells(2, 1).CopyFromRecordset rec
  Else
    arrData = rec.GetRows

    iRec = UBound(arrData, 2) + 1

    For iCol = 0 To rec.Fields.Count - 1
      For iRow = 0 To iRec - 1

        If IsDate(arrData(iCol, iRow)) Then
          arrData(iCol, iRow) = Format(arrData(iCol, iRow))

        ElseIf IsArray(arrData(iCol, iRow)) Then
          arrData(iCol, iRow) = "Array Field"
        End If
      Next iRow
    Next iCol

    ' -- Traspasa los datos a la hoja de Excel
    Hoja.Cells(2, 1).Resize(iRec, rec.Fields.Count).Value = GetData(arrData)
  End If

  Excel.Selection.CurrentRegion.Columns.AutoFit
  Excel.Selection.CurrentRegion.Rows.AutoFit

  ' -- Cierra el recordset y la base de datos y los objetos ADO
  rec.Close
  cn.Close

  Set rec = Nothing
  Set cn = Nothing

  Excel.Visible = True
  'Hoja.PrintPreview

  '        ' -- guardar el libro
  '        Libro.saveAs sOutputPathXLS
  '        Libro.Close

  ' -- Elimina las referencias Xls
  Set Hoja = Nothing
  Set Libro = Nothing
  'Excel.quit
  Set Excel = Nothing

  Exportar_ADO_Excel = True
  Me.Enabled = True
  Exit Function
errSub:
  MsgBox Err.Description, vbCritical, "Error"
  Exportar_ADO_Excel = False
  Me.Enabled = True
End Function

Private Function GetData(vValue As Variant) As Variant
  Dim x As Long, y As Long, xMax As Long, yMax As Long, T As Variant

  xMax = UBound(vValue, 2): yMax = UBound(vValue, 1)

  ReDim T(xMax, yMax)
  For x = 0 To xMax
    For y = 0 To yMax
      T(x, y) = vValue(y, x)
    Next y
  Next x

  GetData = T
End Function


