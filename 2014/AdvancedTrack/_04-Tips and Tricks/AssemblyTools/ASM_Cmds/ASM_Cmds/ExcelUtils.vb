Imports System.IO
Imports Microsoft.Office.Interop

Module ExcelUtils

Public Function IntoStr(ByVal V As Integer) As String
	Return V.ToString("###########0")
End Function

Public Function StrToInt(ByVal V As String) As Integer
	Return Val(IntoStr(V))
End Function

'*********************************************************************************************
'
'  File open utils
'
'*********************************************************************************************

Function IsFileAlreadyOpen(ByRef sName As String) As Boolean
				Dim fs As FileStream
				Try
						fs = File.Open(sName, FileMode.Open, FileAccess.Read, FileShare.None)
						IsFileAlreadyOpen = False
				Catch ex As Exception
						IsFileAlreadyOpen = True
				End Try
End Function

'*********************************************************************************************
'
'  CELL DATA STRING Functions
'
'*********************************************************************************************
Public Function GetCellString(ByRef WS As Excel.Worksheet, ByVal ColRow As String) As String
  Dim strTmp As String
  Dim theCell As Excel.Range
	theCell = WS.Range(ColRow)
	Try
		strTmp = CType(theCell.Text, String)
	Catch ex As Exception
		Return ""
	End Try
	Return strTmp
End Function

Public Sub SetCellString(ByRef WS As Excel.Worksheet, ByVal ColRow As String, ByVal strV As String)
	Dim theCell As Excel.Range
	theCell = WS.Range(ColRow)
	theCell.Value2 = strV
End Sub

Public Function SetCellString_GetNext(ByRef theCell As Excel.Range, ByVal strV As String) As Excel.Range
  Dim NextCell As Excel.Range = theCell.Next
  theCell.Value2 = strV
  Return NextCell
End Function

'*********************************************************************************************
'
'  CELL DATA DATE Functions
'
'*********************************************************************************************
Public Function SetCellDate_GetNext(ByRef theCell As Excel.Range, ByVal theDate As Date) As Excel.Range
	Dim NextCell As Excel.Range = theCell.Next
	theCell.Value2 = theDate.ToString("G")
'  theCell.NumberFormat = "m/d/yyyy h:mm:ss AM/PM"
	theCell.NumberFormat = "m/d/yyyy"
	Return NextCell
End Function

Public Function GetCellDate(ByRef WS As Excel.Worksheet, ByVal ColRow As String) As Date
	Dim aDate As Date
'  Dim tmpStr As String
	Dim theCell As Excel.Range
	theCell = WS.Range(ColRow)

	aDate = DateTime.FromOADate(CType(theCell.Value2, Double))
	Return aDate
End Function

Public Sub SetCellDate(ByRef WS As Excel.Worksheet, ByVal ColRow As String, ByVal theDate As Date)
'  Dim tmpStr As String
	Dim theCell As Excel.Range
	theCell = WS.Range(ColRow)
	theCell.Value2 = theDate.ToString("G")
	theCell.NumberFormat = "m/d/yyyy h:mm:ss AM/PM"
End Sub

'*********************************************************************************************
'
'  CELL DATA INTEGER Functions
'
'*********************************************************************************************

Public Function SetCellInt_GetNext(ByRef theCell As Excel.Range, ByVal N As Integer) As Excel.Range
  Dim NextCell As Excel.Range = theCell.Next
  theCell.Value2 = N
  Return NextCell
End Function

Public Function GetCellInt(ByRef WS As Excel.Worksheet, ByVal ColRow As String) As Integer
	Dim i As Integer
	Dim theCell As Excel.Range
	theCell = WS.Range(ColRow)
	i = CType(theCell.Value2, Integer)
	Return i
End Function

Public Sub SetCellInt(ByRef WS As Excel.Worksheet, ByVal ColRow As String, ByVal v As Integer)
	Dim theCell As Excel.Range
	theCell = WS.Range(ColRow)
	theCell.Value2 = v
End Sub

'*********************************************************************************************
'
'  Misc CELL data Functions
'
'*********************************************************************************************
Public Sub SetColumnTitle(ByRef theRange As Excel.Range, ByVal Title As String)
	theRange.Value2 = Title
	theRange.Font.Bold = True
End Sub


Public Sub SetColumnTitleandWidth(ByRef theRange As Excel.Range, ByVal Title As String, ByVal width As Integer)
  theRange.Value2 = Title
  theRange.ColumnWidth = width
  theRange.Font.Bold = True
End Sub

Public Sub SetCellHlink(ByRef sourceWS As Excel.Worksheet, ByVal sourceColRow As String, ByRef targetWS As Excel.Worksheet, ByVal targetColRow As String)
	Dim theCell As Excel.Range
	Dim addr As String

	theCell = sourceWS.Range(sourceColRow)
	'addr = "File:///" + FileName + " - '" + targetWS.Name + "'!" + targetColRow
	addr = "'" + targetWS.Name + "'!" + targetColRow
	theCell.Hyperlinks.Add(theCell.Item(1), Nothing, addr)
End Sub

End Module
