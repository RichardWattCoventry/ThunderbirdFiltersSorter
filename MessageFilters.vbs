Option Explicit

'*****************************************************************************************
Const FormatASCII = 0
Const ForReading = 1
Const ForWriting = 2

'*****************************************************************************************
Const cstrFileTitle = "Select Thunderbird Message Filters File"
Const cstrFileTypes = "Filter files|*.dat"

'*****************************************************************************************
Const cstrVolatile = "Volatile"
Const cstrSystemRoot = "%SystemRoot%"
Const cstrUserProfile = "%USERPROFILE%"
Const cstrTempPath = "%TEMP%"
Const cstrDocuments = "\Documents"

'*****************************************************************************************
Const cstrOldDelimiter1 = " > "
Const cstrOldDelimiter2 = "> "
Const cstrOldDelimiter3 = " >"
Const cstrNewDelimiter = "/"

'*****************************************************************************************
Dim strDocumentsPath
Dim strFiltersFile
Dim varFiltersList
Dim varFilterNames

'*****************************************************************************************
strDocumentsPath = GetDocumentsPath
If (WScript.Arguments.Count > 0) Then
	strFiltersFile = WScript.Arguments.Item(0)
Else
	strFiltersFile = SelectFile(cstrFileTitle, strDocumentsPath, cstrFileTypes)
End If
If (strFiltersFile <> vbNullString) Then
	varFiltersList = ReadFiltersList(strFiltersFile)
	If Not IsEmpty(varFiltersList) Then
		varFilterNames = IndexFilterNames(varFiltersList)
		If Not IsEmpty(varFilterNames) Then
			'Call ReplaceDelimiters(varFilterNames, cstrOldDelimiter1, cstrNewDelimiter)
			'Call ReplaceDelimiters(varFilterNames, cstrOldDelimiter2, cstrNewDelimiter)
			'Call ReplaceDelimiters(varFilterNames, cstrOldDelimiter3, cstrNewDelimiter)
			'Call OutputArray(varFilterNames, ".\NamesBefore.txt")
			Call QuickSort(varFilterNames, LBound(varFilterNames, 1), UBound(varFilterNames, 1), 0)
			'Call OutputArray(varFilterNames, ".\NamesAfter.txt")
			varFiltersList = RebuildFiltersList(varFiltersList, varFilterNames)
			If Not IsEmpty(varFiltersList) Then
				Call WriteFiltersList(strFiltersFile, varFiltersList)
				Erase varFilterNames
				Erase varFiltersList
			End If
		End If	
	End If
	strFiltersFile = vbNullString
End If
strDocumentsPath = vbNullString
WScript.Quit

'*****************************************************************************************
'*	Script functions and subs:

'#########################################################################################
'#	Function GetDocumentsPath():
'#	This function determines the "My Documents" folder path for the current user from the
'#	environment variables and returns the path.
Function GetDocumentsPath
	Dim strPath
	Dim objShell
	
	strPath = vbNullString
	Set objShell = CreateObject("WScript.Shell")
	If Not (objShell is Nothing) Then
		strPath = objShell.ExpandEnvironmentStrings(cstrUserProfile)
		If (strPath <> vbNullString) Then
			strPath = strPath & cstrDocuments
		End If
		Set objShell = Nothing
	End If
	GetDocumentsPath = strPath
End Function

'#########################################################################################
'#	Function GetTempPath():
'#	This function determines the temporary folder path for the current user from the
'#	environment variables and returns the path.
Function GetTempPath
	Dim strPath
	Dim strSysRoot
	Dim objShell
	
	strPath = vbNullString
	Set objShell = CreateObject("WScript.Shell")
	If Not (objShell is Nothing) Then
		strSysRoot = objShell.ExpandEnvironmentStrings(cstrSystemRoot)
		strPath = objShell.ExpandEnvironmentStrings(cstrTempPath)
		If (strPath <> vbNullString) Then 
			strPath = strPath
		End If
		If (strSysRoot <> vbNullString) Then 
			strPath = Replace(strPath, cstrSystemRoot, strSysRoot)
		End If
		Set objShell = Nothing
	End If
	GetTempPath = strPath
End Function

'#########################################################################################
'#	Function SelectFile(TitleText, InPath, FileTypes):
'#	  TitleText: String - dialog box title text
'#	  InPath   : String - folder path
'#	  FileTypes: String - file types list
'#	This function .
Function SelectFile(TitleText, InPath, FileTypes)
	Dim strFileName
	Dim objSelector
	
	strFileName = vbNullString
	TitleText = Trim(TitleText)
	InPath = Trim(InPath)
	FileTypes = Trim(FileTypes)
	If ((TitleText <> vbNullString) And (FileTypes <> vbNullString)) Then
		Set objSelector = CreateObject("WshKit.Browse")
		If Not (objSelector Is Nothing) Then
			With objSelector
				.Title = TitleText
				.Filter = FileTypes
				strFileName = .Browse(InPath)
				If (.FileCount = 0) Then strFileName = vbNullString
			End With
			Set objSelector = Nothing
		End If
	End If
	SelectFile = strFileName
End Function

'#########################################################################################
'#	Function ReadFiltersList(FiltersFile):
'#	  FiltersFile: String - filters filename
'#	This function .
Function ReadFiltersList(FiltersFile)
	Dim varFilters
	Dim objFSO
	Dim objStream
	Dim strText
	
	varFilters = Empty
	FiltersFile = Trim(FiltersFile)
	If (FiltersFile <> vbNullString) Then
		Set objFSO = CreateObject("Scripting.FileSystemObject")
		If Not (objFSO Is Nothing) Then
			If objFSO.FileExists(FiltersFile) Then
				Set objStream = objFSO.OpenTextFile(FiltersFile, ForReading, False)
				If Not (objStream Is Nothing) Then
					With objStream
						strText = .ReadAll
						.Close
					End With
					Set objStream = Nothing
					varFilters = Split(strText, vbCrLf)
				End If
			End If
			Set objFSO = Nothing
			strText = vbNullString
		End If
	End If
	ReadFiltersList = varFilters
End Function

'#########################################################################################
'#	Function IndexFilterNames(FiltersList):
'#	  FiltersList: Variant - filters list
'#	This function .
Function IndexFilterNames(FiltersList)
	Dim avarIndex
	Dim lngLine
	Dim lngIndex
	Dim strLine
	
	avarIndex = Empty
	If Not IsEmpty(FiltersList) Then
		ReDim avarIndex(UBound(FiltersList), 2)
		lngIndex = LBound(avarIndex, 1) - 1
		For lngLine = LBound(FiltersList) To UBound(FiltersList)
			strLine = Trim(FiltersList(lngLine))
			If (InStr(1, strLine, "name=") > 0) Then
				lngIndex = lngIndex + 1
				avarIndex(lngIndex, 0) = strLine
				avarIndex(lngIndex, 1) = lngLine
				avarIndex(lngIndex, 2) = 0
			End If
		Next
		If (lngIndex < UBound(avarIndex, 1)) Then
			avarIndex = TransposeArray(avarIndex)
			ReDim Preserve avarIndex(2, lngIndex)
			avarIndex = TransposeArray(avarIndex)
		End If
		For lngIndex = LBound(avarIndex, 1) To UBound(avarIndex, 1)
			If (lngIndex < UBound(avarIndex, 1)) Then
				avarIndex(lngIndex, 2) = (avarIndex((lngIndex + 1), 1) - 1)
			Else
				avarIndex(lngIndex, 2) = UBound(FiltersList)
			End If
		Next
	End If
	IndexFilterNames = avarIndex
End Function

'#########################################################################################
'#	Sub ReplaceDelimiters(FiltersList, OldDelimiter, NewDelimiter):
'#	  FiltersList : Variant - filters list
'#	  OldDelimiter: String  - old delimiter
'#	  NewDelimiter: String  - new delimiter
'#	This sub .
Sub ReplaceDelimiters(FiltersList, OldDelimiter, NewDelimiter)
	Dim lngIndex
	Dim strName
	
	OldDelimiter = Trim(OldDelimiter)
	If Not (IsEmpty(FiltersList) Or (OldDelimiter = vbNullString) Or (NewDelimiter = vbNullString)) Then
		For lngIndex = LBound(FiltersList, 1) To uBound(FiltersList, 1)
			strName = Trim(FiltersList(lngIndex, 1))
			If (InStr(1, strName, OldDelimiter) > 0) Then
				strName = Replace(strName, OldDelimiter, NewDelimiter)
				FiltersList(lngIndex, 1) = strName
			End If
		Next
	End If
End Sub

'#########################################################################################
'#	Function RebuildFiltersList(FiltersList, FiltersIndex):
'#	  FiltersList : Variant - filters list array
'#	  FiltersIndex: Variant - filter index array
'#	This function .
Function RebuildFiltersList(FiltersList, FiltersIndex)
	Dim avarNewList
	Dim lngIndex
	Dim lngNewLine
	Dim lngRow1
	Dim lngRow2
	Dim avarFilter
	
	avarNewList = Empty
	If Not (IsEmpty(FiltersList) Or IsEmpty(FiltersIndex)) Then
		ReDim avarNewList(UBound(FiltersList))
		lngNewLine = LBound(avarNewList)
		avarNewList(lngNewLine) = FiltersList(lngNewLine)
		lngNewLine = lngNewLine + 1
		avarNewList(lngNewLine) = FiltersList(lngNewLine)
		lngNewLine = lngNewLine + 1
		For lngIndex = LBound(FiltersIndex, 1) To UBound(FiltersIndex, 1)
			lngRow1 = FiltersIndex(lngIndex, 1)
			lngRow2 = FiltersIndex(lngIndex, 2)
			avarFilter = ReadFilter(FiltersList, lngRow1, lngRow2)
			If Not IsEmpty(avarFilter) Then
				Call WriteFilter(avarNewList, avarFilter, lngNewLine)
				Erase avarFilter
			End If
		Next
		Erase FiltersList
	End If
	RebuildFiltersList = avarNewList
End Function

'#########################################################################################
'#	Function ReadFilter(FiltersList, StartRow, EndRow):
'#	  FiltersList: Variant - filters list array
'#	  StartRow   : Long    - filter start row
'#	  EndRow     : Long    - filter end row
'#	This function .
Function ReadFilter(FiltersList, StartRow, EndRow)
	Dim avarFilter
	Dim lngSize
	Dim lngLine
	Dim lngIndex
	
	avarFilter = Empty
	If Not IsEmpty(FiltersList) Then
		lngSize = (EndRow - StartRow)
		ReDim avarFilter(lngSize)
		lngIndex = LBound(avarFilter)
		For lngLine = StartRow To EndRow
			avarFilter(lngIndex) = FiltersList(lngLine)
			lngIndex = lngIndex + 1
		Next
	End If
	ReadFilter = avarFilter
End Function

'#########################################################################################
'#	Sub WriteFilter(NewFiltersList, NewFilter, StartLine):
'#	  NewFiltersList: Variant - new filters list array
'#	  NewFilter     : Variant - new filter array
'#	  StartLine     : Long    - new filter start line
'#	This sub .
Sub WriteFilter(NewFiltersList, NewFilter, StartLine)
	Dim lngNewLine
	
	If Not (IsEmpty(NewFiltersList) Or IsEmpty(NewFilter)) Then
		For lngNewLine = LBound(NewFilter) To UBound(NewFilter)
			NewFiltersList(StartLine) = NewFilter(lngNewLine)
			StartLine = StartLine + 1
		Next
	End If
End Sub

'#########################################################################################
'#	Sub WriteFiltersList(FiltersFile, FiltersList):
'#	  FiltersFile: String  - filters filename
'#	  FiltersList: Variant - filters list array
'#	This sub .
Sub WriteFiltersList(FiltersFile, FiltersList)
	Dim objFSO
	Dim objStream
	Dim strText
	
	FiltersFile = Trim(FiltersFile)
	If Not (IsEmpty(FiltersList) Or (FiltersFile = vbNullString)) Then
		Set objFSO = CreateObject("Scripting.FileSystemObject")
		If Not (objFSO Is Nothing) Then
			If objFSO.FileExists(FiltersFile) Then
				Set objStream = objFSO.OpenTextFile(FiltersFile, ForWriting, True)
				If Not (objStream Is Nothing) Then
					strText = Join(FiltersList, vbCrLf)
					With objStream
						.Write strText
						.Close
					End With
					Set objStream = Nothing
				End If
			End If
			Set objFSO = Nothing
		End If
	End If
End Sub

'*****************************************************************************************
'*	Supporting functions and subs:

'#########################################################################################
'#	Function TransposeArray(DataArray):
'#	  DataArray: Variant - data array
'#	This function .
Function TransposeArray(DataArray)
	Dim varTemp
	Dim lngLow1
	Dim lngUpp1
	Dim lngLow2
	Dim lngUpp2
	Dim lngI
	Dim lngJ
	
	varTemp = Empty
	If Not IsEmpty(DataArray) Then
		lngLow1 = LBound(DataArray, 1)
		lngUpp1 = UBound(DataArray, 1)
		lngLow2 = LBound(DataArray, 2)
		lngUpp2 = UBound(DataArray, 2)
		ReDim varTemp(lngUpp2, lngUpp1)
		For lngI = lngLow1 To lngUpp1
			For lngJ = lngLow2 To lngUpp2
				varTemp(lngJ, lngI) = DataArray(lngI, lngJ)
			Next
		Next
	End If
	TransposeArray = varTemp
End Function

'#########################################################################################
'#	Sub SwapRows(DataArray, Row1, Row2):
'#	  DataArray: Variant - data array
'#	  Row1     : Long    - first row
'#	  Row2     : Long    - second row
'#	This sub .
Sub SwapRows(DataArray, Row1, Row2)
	Dim lngX
	Dim varTemp
	
	For lngX = LBound(DataArray, 2) To UBound(DataArray, 2)
		varTemp = DataArray(Row1, lngX)
		DataArray(Row1, lngX) = DataArray(Row2, lngX)
		DataArray(Row2, lngX) = varTemp
	Next
End Sub

'#########################################################################################
'#	Sub QuickSort(DataArray, LowBound, HighBound, SortField):
'#	  DataArray: Variant - data array
'#	  LowBound : Long    - low array bound
'#	  HighBound: Long    - high array bound
'#	  SortField: Long    - sort field number
'#	This sub .
Sub QuickSort(DataArray, LowBound, HighBound, SortField)
	Dim arrPivot()
	Dim lngLoSwap
	Dim lngHiSwap
	Dim varTemp
	Dim lngCounter
	
	ReDim arrPivot(UBound(DataArray, 2))
	If ((HighBound - LowBound) = 1) Then
		If DataArray(LowBound, SortField) > DataArray(HighBound, SortField) Then
			Call SwapRows(DataArray, HighBound, LowBound)
		End If
	End If
	For lngCounter = 0 To UBound(DataArray,2 )
		arrPivot(lngCounter) = DataArray(Int((LowBound + HighBound) / 2), lngCounter)
		DataArray(Int((LowBound + HighBound) / 2), lngCounter) = DataArray(LowBound, lngCounter)
		DataArray(lowBound, lngCounter) = arrPivot(lngCounter)
	Next
	lngLoSwap = LowBound + 1
	lngHiSwap = HighBound
	Do	
		While (lngLoSwap < lngHiSwap) And (StrComp(DataArray(lngLoSwap, SortField), arrPivot(SortField), 1) < 0)
			lngLoSwap = lngLoSwap + 1
		Wend
		While (StrComp(DataArray(lngHiSwap, SortField), arrPivot(SortField), 1) > 0)
			lngHiSwap = lngHiSwap - 1
		Wend
		If (lngloswap < lngHiSwap) Then Call SwapRows(DataArray, lngLoSwap, lngHiSwap)
	Loop While (lngLoSwap < lngHiSwap)
	For lngCounter = LBound(DataArray, 2) To UBound(DataArray, 2)
		DataArray(lowbound, lngCounter) = DataArray(lngHiSwap, lngCounter)
		DataArray(lngHiSwap, lngCounter) = arrPivot(lngCounter)
	Next
	If (LowBound < (lngHiSwap - 1)) Then Call QuickSort(DataArray, LowBound, (lngHiSwap - 1), SortField)
	If ((lngHiSwap + 1) < HighBound) Then Call QuickSort(DataArray, (lngHiSwap + 1), HighBound, SortField)
End Sub

'#########################################################################################
'#	Sub OutputArray(DataArray, ArrayFile):
'#	  DataArray: Variable - data array
'#	  ArrayFile: String   - filename
'#	This sub .
Sub OutputArray(DataArray, ArrayFile)
	Dim objFSO
	Dim objStream
	Dim lngRow
	Dim lngCol
	Dim strItem
	
	ArrayFile = Trim(ArrayFile)
	If Not (IsEmpty(DataArray) Or (ArrayFile = vbNullString)) Then
		Set objFSO = CreateObject("Scripting.FileSystemObject")
		If Not (objFSO Is Nothing) Then
			Set objStream = objFSO.OpenTextFile(ArrayFile, ForWriting, True)
			If Not (objStream Is Nothing) Then
				For lngRow = LBound(DataArray, 1) To UBound(DataArray, 1)
					For lngCol = LBound(DataArray, 2) To UBound(DataArray, 2)
						strItem = "(" & lngRow & "," & lngCol & "): "
						strItem = strItem & DataArray(lngRow, lngCol) & vbCrLf
						objStream.Write strItem
					Next
				Next
				objStream.Close
				Set objStream = Nothing
			End If
			Set objFSO = Nothing
		End If
	End If
End Sub
