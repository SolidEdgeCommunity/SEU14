Imports System.Reflection
Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Excel

Public Class clsExcelData
	'Create the new Excel file
	Private excel_app As Excel.Application = Nothing
	Private xlsNameFile As Excel.Workbook = Nothing
	Private theWS As Excel.Worksheet = Nothing
	Private m_InitOK As Boolean = False

	Dim m_fdir As String = "C:\Users\aspatric\Local_Documents\SEU13\00_Developer\_CopyRenameReplace\"
	Dim m_fname As String = "file_names.xlsx"

	Dim strDateFormat = "MM'/'dd'/'yyyy"

	Public Sub InitClass()
	Dim wb As Excel.Workbooks
	Dim obj As Object

	'Open Excel and get the worksheets
  '
  'NOTE: We want to always use the xls in the release folder so replace with copy
  'names are update to date when running either debug or release.

  Dim pExecutingAssembly As System.Reflection.Assembly
	pExecutingAssembly = System.Reflection.Assembly.GetExecutingAssembly()
	Dim tmpStr As String = pExecutingAssembly.Location

  If tmpStr.ToLower.Contains("debug") Then
    'Replace "Debug" with release.
    tmpStr = tmpStr.Substring(0, tmpStr.LastIndexOf("D"))
    m_fdir = tmpStr + "Release"
  Else
    'This is the release or addin install folder.
    m_fdir = tmpStr.Substring(0, tmpStr.LastIndexOf("\"))
  End If

	Try
		Try
				'See if Excel is already running
				excel_app = GetObject(, "Excel.Application")
				If Not IsNothing(excel_app) Then
					'See if our doc is open and close it.

					Dim xlsFile As Excel.Workbook = Nothing
					 For Each xlsFile In excel_app.Workbooks
						If xlsFile.Name = m_fname Then
							xlsFile.Close(True)
						End If
					 Next
				End If
		Catch ex As Exception
      'Excel was not running continue
		End Try

		'Start our own instance of Excel
		excel_app = New Excel.Application

		obj = excel_app.Workbooks
		wb = CType(obj, Excel.Workbooks)

		xlsNameFile = wb.Open(m_fdir + "\" + m_fname)

	Catch ex As Exception
			MsgBox(ex.Message, , "clsExcelData-InitClass")
			m_InitOK = False
			Exit Sub
	End Try
	m_InitOK = True

	End Sub

	Public Sub ShutDown()
	'Close down Excel
	Try
		theWS = Nothing
		If Not IsNothing(xlsNameFile) Then
      xlsNameFile.Close(False)
		End If

		excel_app.Quit()
		excel_app = Nothing

	Catch ex As Exception
			MsgBox(ex.Message, , "clsExcelData-Finalize")
	End Try

	End Sub

	Property FileName() As String
		Get
			Return m_fname
		End Get
		Set(ByVal value As String)
			m_fname = value
		End Set
	End Property

	ReadOnly Property InitOK() As Boolean
		Get
			Return m_InitOK
		End Get
	End Property


Public Sub SaveXls()
			xlsNameFile.Save()
End Sub

Public Function GetNewFileName(ByVal strCurrentName As String, ByVal DocType As SolidEdgeFramework.DocumentTypeConstants) As String

	Dim strNewFileName As String = ""
	Dim ii As Integer = 0

	strNewFileName = GenerateNewFileNameString(strCurrentName, DocType)
	xlsNameFile.Save()
	Return strNewFileName

End Function

Private Function GenerateNewFileNameString(ByVal strCurrentName As String, ByVal DocType As SolidEdgeFramework.DocumentTypeConstants) As String

	Dim strFileNameNoExt As String
	Dim strNewFileName As String = ""
	Dim fileWS As Excel.Worksheet = Nothing
	Dim dashIdx As Integer
	Dim ii As Integer
	Dim tmpStr As String
	Dim bMatch As Boolean = False
	Dim currIdx As Integer

	'Strip out the dash number (if present)
	dashIdx = strCurrentName.LastIndexOf("-")

	If dashIdx > -1 Then
		strFileNameNoExt = strCurrentName.Substring(0, dashIdx)
	Else
		strFileNameNoExt = strCurrentName.Substring(0, strCurrentName.Length - 4)
	End If

	Select Case DocType

		Case SolidEdgeFramework.DocumentTypeConstants.igAssemblyDocument

			fileWS = xlsNameFile.Worksheets(1)

		Case SolidEdgeFramework.DocumentTypeConstants.igPartDocument

			fileWS = xlsNameFile.Worksheets(2)

		Case SolidEdgeFramework.DocumentTypeConstants.igSheetMetalDocument
			fileWS = xlsNameFile.Worksheets(3)
	End Select

	'See if the name is already in the WS if not add it with index 1.
	'If present increment the index in the WS and add the index to the string

	ii = 2
	bMatch = False

	While True
		tmpStr = GetCellString(fileWS, "A" & IntoStr(ii))

		If strFileNameNoExt.ToLower = tmpStr.ToLower Then
			bMatch = True
			Exit While
		ElseIf tmpStr = "" Then
			bMatch = False
			Exit While
		End If
		ii = ii + 1
	End While

	If bMatch Then
		'Get the number
		currIdx = GetCellInt(fileWS, "B" & IntoStr(ii))
		currIdx = currIdx + 1

		strNewFileName = strFileNameNoExt + "-" + Format(currIdx, "00000")

		'Update the number
		SetCellInt(fileWS, "B" & IntoStr(ii), currIdx)

	Else
		'Add the filename and set the number to 1
		currIdx = 1

		SetCellString(fileWS, "A" & IntoStr(ii), strFileNameNoExt)

		strNewFileName = strFileNameNoExt + "-" + Format(currIdx, "00000")

		'Update the number
		SetCellInt(fileWS, "B" & IntoStr(ii), currIdx)
	End If

	Select Case DocType

		Case SolidEdgeFramework.DocumentTypeConstants.igAssemblyDocument

			strNewFileName = strNewFileName + ".asm"

		Case SolidEdgeFramework.DocumentTypeConstants.igPartDocument

			strNewFileName = strNewFileName + ".par"

		Case SolidEdgeFramework.DocumentTypeConstants.igSheetMetalDocument
			strNewFileName = strNewFileName + ".psm"
	End Select

	Return strNewFileName
End Function

End Class
