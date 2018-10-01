Option Explicit

' Dependencies:
'   - cCustomErrors

' Private Variables
Private vHRng As Range

' Public Set Properties
Property Set HeaderRange(ByRef rng As Range)
    ' Description: Define the header range
    ' Dependencies: None
    ' Inputs:
    '   - (As Range)    rng - The range of cells to be recorded as a header row
    ' Outputs: None
    On Error GoTo ErrorHandle
    Set vHRng = rng
    Exit Property
ErrorHandle:
    custErr.RaiseError "cHeaderRange - Set HeaderRange()"
End Property

' Public Get Properties
Property Get HeaderRange() As Range
    ' Description: Return the header range as an object
    ' Dependencies:
    '   - The range must have been defined
    ' Inputs: None
    ' Outputs:
    '   - (As Range) The range of cells in the header range
    On Error GoTo ErrorHandle
    If vHRng Is Nothing Then Err.Raise custErr.GenericError, Description:="The header range has not been defined"
    Set HeaderRange = vHRng
    Exit Property
ErrorHandle:
    custErr.RaiseError "cHeaderRange - Get HeaderRange()"
End Property


Property Get SheetRow() As Long
    ' Description: Return the sheet row that the header is in
    ' Dependencies:
    '   - Me.HeaderRange()
    ' Inputs: None
    ' Outputs:
    '   - (As Long) The sheet row that the header range is in
    On Error GoTo ErrorHandle
    SheetRow = Me.HeaderRange.Row
    Exit Property
ErrorHandle:
    custErr.RaiseError "cHeaderRange - Get SheetRow()"
End Property


Property Get ColPresent(ByVal ColName As String) As Boolean
    ' Description: Return whether the named column exists or not
    ' Dependencies:
    '   - Me.HeaderRange()
    ' Inputs:
    '   - (As String)   ColName     - The name of a column to be found
    ' Outputs:
    '   - (As Boolean) Whether the given column name is present in the header or not
    On Error GoTo ErrorHandle
    ColPresent = Not Me.HeaderRange.Find(what:=ColName, LookIn:=XlFindLookIn.xlValues, LookAt:=XlLookAt.xlWhole, MatchCase:=True) Is Nothing
    Exit Property
ErrorHandle:
    custErr.RaiseError "cHeaderRange - Get ColPresent()"
End Property

Property Get ColCell(ByVal ColName As String) As Range
    ' Description: Return the cell in the header with the column name given
    ' Dependencies:
    '   - Me.ColPresent()
    '   - Me.HeaderRange()
    ' Inputs:
    '   - (As String)   ColName     - The name of a column to be found
    ' Outputs:
    '   - (As Range) The cell in the header with the column name
    On Error GoTo ErrorHandle
    If Not Me.ColPresent(ColName) Then Err.Raise custErr.GenericError, Description:="The given column name, " & ColName & ", could not be found in the header"
    Set ColCell = Me.HeaderRange.Find(what:=ColName, LookIn:=XlFindLookIn.xlValues, LookAt:=XlLookAt.xlWhole, MatchCase:=True)
    Exit Property
ErrorHandle:
    custErr.RaiseError "cHeaderRange - Get ColCell()"
End Property

Property Get ColInSheetNum(ByVal ColName As String) As Long
    ' Description: Return the sheet column number that the given header name is in
    ' Dependencies:
    '   - Me.ColCell()
    ' Inputs:
    '   - (As String)   ColName     - The name of a column to be found
    ' Outputs:
    '   - (As Long) The sheet column number that the defined column header is in
    On Error GoTo ErrorHandle
    ColInSheetNum = Me.ColCell.Column
    Exit Property
ErrorHandle:
    custErr.RaiseError "cHeaderRange - Get ColInSheetNum()"
End Property

Property Get ColInSheetLetter(ByVal ColName As String) As String
    ' Description: Return the sheet column letter that the given header name is in
    ' Dependencies:
    '   - Me.ColInSheetNum()
    ' Inputs:
    '   - (As String)   ColName     - The name of a column to be found
    ' Outputs:
    '   - (As String) The sheet column letter that the defined column header is in
    On Error GoTo ErrorHandle
    ColInSheetLetter = Replace(Cells(1, Me.ColInSheetNum(ColName)).Address(RowAbsolute:=False, ColumnAbsolute:=False), 1, "")
    Exit Property
ErrorHandle:
    custErr.RaiseError "cHeaderRange - Get ColInSheetLetter()"
End Property

Property Get ColInData(ByVal ColName As String) As Long
    ' Description: Return the data column number that the given header name is in
    ' Dependencies:
    '   - Me.ColInSheetNum()
    '   - Me.HeaderRange()
    ' Inputs:
    '   - (As String)   ColName     - The name of a column to be found
    ' Outputs:
    '   - (As Long) The data column number that the defined column header is in
    On Error GoTo ErrorHandle
    ColInData = Me.ColInSheetNum(ColName) - Me.HeaderRange.Columns(1).Column + 1
    Exit Property
ErrorHandle:
    custErr.RaiseError "cHeaderRange - Get ColInData()"
End Property

Property Get WholeSheetCol(ByValColName As String) As Range
    ' Description: Return the whole sheet column that the defined column header is in
    ' Dependencies:
    '   - Me.ColCell()
    ' Inputs:
    '   - (As String)   ColName     - The name of a column to be found
    ' Outputs:
    '   - (As Range) The whole sheet column that the defined column header is in
    On Error GoTo ErrorHandle
    WholeSheetCol = Me.ColCell.EntireColumn
    Exit Property
ErrorHandle:
    custErr.RaiseError "cHeaderRange - Get WholeSheetCol()"
End Property
