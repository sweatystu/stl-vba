Option Explicit

' Dependencies:
'   - cCustomErrors class
'   - cHeaderRange class
'   - cDataRange class

' Private Variables
Private vRng As Range
Private vHeaderRange As cHeaderRange
Private vDataRange As cDataRange

' Public Procedures
Sub DefineRange(ByRef CellRange As Range, Optional HasHeaderRow As Boolean = True, Optional HeaderRow As Long = 1)
    ' Description: Defines the range and initialises the header row and data range
    ' Dependencies: None
    ' Inputs:
    '   - (As Range)        CellRange       - Range of cells defining the whole range
    '   - (As Boolean, opt) HasHeaderRow    - Define whether the range has a header row or not
    '   - (As Long, opt)    HeaderRow       - The row in the data (not in the sheet) that contains the header row
    '                                       - Must be greater than 0 and must be smaller than the number of rows in the whole data range
    ' Outputs: None
    On Error GoTo ErrorHandle
    Set vRng = CellRange
    Set vDataRange = New cDataRange
    If HasHeaderRow Then
        If HeaderRow < 1 Then Err.Raise custErr.GenericError, Description:="The header-row number (" & HeaderRow & ") must be greater than 0"
        If HeaderRow >= vRng.Rows.Count Then Err.Raise custErr.GenericError, Description:="The header-row number (" & HeaderRow & ") must be less than the number of rows in the data (" & vRng.Rows.Count & ")"
        Set vHeaderRange = New cHeaderRange
        Set vHeaderRange.HeaderRange = vRng.Rows(HeaderRow)
        Set vDataRange.DataRange = Intersect(vRng, vRng.Parent.Rows(vHeaderRange.SheetRow + 1 & ":" & vRng.Rows(vRng.Rows.Count).Row))
    Else
        Set vHeaderRange = Nothing
        Set vDataRange = vRng
    End If
    Exit Sub
ErrorHandle:
    custErr.RaiseError "cRange - DefineRange()"
End Sub

' Public Get Properties
Property Get rng() As Range
    ' Description: Object containing the whole data range
    ' Dependencies:
    '   - Range must have been defined via DefineRange()
    ' Inputs: None
    ' Outputs:
    '   - (As Range) Range of cells defining the whole range
    On Error GoTo ErrorHandle
    If vRng Is Nothing Then Err.Raise custErr.GenericError, Description:="The range has not been defined"
    Set rng = vRng
    Exit Property
ErrorHandle:
    custErr.RaiseError "cRange - Get rng()"
End Property

Property Get Sheet() As Worksheet
    ' Description: The worksheet that the defined range is in
    ' Dependencies:
    '   - Me.rng()
    ' Inputs: None
    ' Outputs:
    '   - (As Worksheet) Worksheet the defined range is in
    On Error GoTo ErrorHandle
    Set Sheet = Me.rng.Parent
    Exit Property
ErrorHandle:
    custErr.RaiseError "cRange - Get sheet()"
End Property

Property Get wb() As Workbook
    ' Description: The workbook that the defined range is in
    ' Dependencies:
    '   - Me.Sheet()
    ' Inputs: None
    ' Outputs:
    '   - (As Workbook) Workbook the defined range is in
    On Error GoTo ErrorHandle
    Set wb = Me.Sheet.Parent
    Exit Property
ErrorHandle:
    custErr.RaiseError "cRange - Get sheet()"
End Property

Property Get HeaderRange() As cHeaderRange
    ' Description: Provides access to the properties of the header range
    ' Dependencies:
    '   - Range must have been defined via DefineRange() with HasHeader set to True
    '   - cHeaderRange class
    ' Inputs: None
    ' Outputs:
    '   - (As cHeaderRange) Header range
    On Error GoTo ErrorHandle
    If vRng Is Nothing Then Err.Raise custErr.GenericError, Description:="The range has not been defined"
    If vHeaderRange Is Nothing Then Err.Raise custErr.GenericError, Description:="The range does not have a header"
    Set HeaderRange = vHeaderRange
    Exit Property
ErrorHandle:
    custErr.RaiseError "cRange - Get HeaderRange()"
End Property

Property Get DataRange() As cDataRange
    ' Description: Provides access to the properties of the data range
    ' Dependencies:
    '   - Range must have been defined via DefineRange()
    '   - cDataRange class
    ' Inputs: None
    ' Outputs:
    '   - (As cDataRange) Data range below the header
    On Error GoTo ErrorHandle
    If vRng Is Nothing Then Err.Raise custErr.GenericError, Description:="The range has not been defined"
    If vDataRange Is Nothing Then Err.Raise custErr.GenericError, Description:="The data range has not been defined"
    Set DataRange = vDataRange
    Exit Property
ErrorHandle:
    custErr.RaiseError "cRange - Get DataRange()"
End Property

Property Get NoColumns() As Long
    ' Description: The number of columns in the range
    ' Dependencies:
    '   - Me.rng()
    ' Inputs: None
    ' Outputs:
    '   - (As Long) The number of columns
    On Error GoTo ErrorHandle
    NoColumns = Me.rng.Columns.Count
    Exit Property
ErrorHandle:
    custErr.RaiseError "cRange - Get NoColumns()"
End Property

Property Get FirstColNum() As Long
    ' Description: The sheet column number of the first column in the range
    ' Dependencies:
    '   - Me.rng()
    ' Inputs: None
    ' Outputs:
    '   - (As Long) The sheet number of the first column
    On Error GoTo ErrorHandle
    FirstCol = Me.rng.Columns(1).Column
    Exit Property
ErrorHandle:
    custErr.RaiseError "cRange - Get FirstCol()"
End Property

Property Get LastColNum() As Long
    ' Description: The sheet column number of the last column in the range
    ' Dependencies:
    '   - Me.rng()
    ' Inputs: None
    ' Outputs:
    '   - (As Long) The sheet number of the last column
    On Error GoTo ErrorHandle
    LastCol = Me.rng.Columns(Me.rng.Columns.Count).Column
    Exit Property
ErrorHandle:
    custErr.RaiseError "cRange - Get LastCol()"
End Property

Property Get FirstColLetter() As String
    ' Description: The sheet column letter of the first column in the range
    ' Dependencies:
    '   - Me.FirstColNum()
    ' Inputs: None
    ' Outputs:
    '   - (As String) The sheet letter of the first column
    On Error GoTo ErrorHandle
    FirstColLetter = Replace(Cells(1, Me.FirstColNum).Address(RowAbsolute:=False, ColumnAbsolute:=False), 1, "")
    Exit Property
ErrorHandle:
    custErr.RaiseError "cRange - Get FirstColLetter()"
End Property

Property Get LastColLetter() As String
    ' Description: The sheet column letter of the last column in the range
    ' Dependencies:
    '   - Me.LastColNum()
    ' Inputs: None
    ' Outputs:
    '   - (As String) The sheet letter of the last column
    On Error GoTo ErrorHandle
    LastColLetter = Replace(Cells(1, Me.LastColNum).Address(RowAbsolute:=False, ColumnAbsolute:=False), 1, "")
    Exit Property
ErrorHandle:
    custErr.RaiseError "cRange - Get LastColLetter()"
End Property

Property Get Cell_DataRowColName(ByVal DataRow As Long, ByVal ColName As String) As Range
    ' Description: The cell in the x^th row of data, in the defined column
    ' Dependencies:
    '   - Me.DataRange()
    '   - Me.HeaderRange()
    ' Inputs:
    '   - (As Long)     DataRow - The numbered row in the data from which to return a result from
    '   - (As String)   ColName - The name of the column from which to return a result from
    ' Outputs:
    '   - (As Range) The cell in the x^th row of data, in the defined column
    Dim result As Range
    On Error GoTo ErrorHandle
    Set result = Intersect(Me.DataRange.RowInData(DataRow), Me.HeaderRange.WholeSheetCol(ColName))
    If result Is Nothing Then Err.Raise custErr.GenericError, Description:="The row (" & DataRow & ") and column (" & ColName & ") combination has not been recognised"
    Set Cell_DataRowColName = result
    Exit Property
ErrorHandle:
    custErr.RaiseError "cRange - Get Cell_DataRowColName()"
End Property

Property Get Cell_SheetRowColName(ByVal SheetRow As Long, ByVal ColName As String) As Range
    ' Description: The cell in the x^th sheet row, in the defined column
    ' Dependencies:
    '   - Me.DataRange()
    '   - Me.HeaderRange()
    ' Inputs:
    '   - (As Long)     SheetRow    - The numbered row in the sheet from which to return a result from
    '   - (As String)   ColName     - The name of the column from which to return a result from
    ' Outputs:
    '   - (As Range) The cell in the x^th sheet row, in the defined column
    Dim result As Range
    On Error GoTo ErrorHandle
    Set result = Intersect(Me.Sheet.Rows(SheetRow), Me.HeaderRange.WholeSheetCol(ColName))
    If result Is Nothing Then Err.Raise custErr.GenericError, Description:="The sheet row (" & SheetRow & ") and column (" & ColName & ") combination has not been recognised"
    Set Cell_SheetRowColName = result
    Exit Property
ErrorHandle:
    custErr.RaiseError "cRange - Get Cell_SheetRowColName()"
End Property

Property Get Cell_UnderHeader(ByVal RowNum As Long, ByVal ColName As String) As Range
    ' Description: The cell in the x^th row underneath the header row, in the defined column
    ' Dependencies:
    '   - Me.HeaderRange()
    '   - Me.DataRange()
    ' Inputs:
    '   - (As Long)     RowNum      - The number of the row underneath the header from which to return a result from
    '                               - Must be greater than 0
    '   - (As String)   ColName     - The name of the column from which to return a result from
    ' Outputs:
    '   - (As Range) The cell in the x^th row beneath the header row, in the defined column
    Dim result As Range
    On Error GoTo ErrorHandle
    If RowNum < 1 Then Err.Raise custErr.GenericError, Description:="The row number given (" & RowNum & ") should be greater than 0"
    Set result = Me.HeaderRange.ColCell(ColName).Offset(RowNum, 0)
    If result Is Nothing Then Err.Raise custErr.GenericError, Description:="The sheet row (" & SheetRow & ") and column (" & ColName & ") combination has not been recognised"
    Set Cell_UnderHeader = result
    Exit Property
ErrorHandle:
    custErr.RaiseError "cRange - Get Cell_UnderHeader()"
End Property
