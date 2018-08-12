Option Explicit

Private vHRng As Range


Property Set HeaderRange(ByRef rng As Range)
    On Error GoTo ErrorHandle
    Set vHRng = rng
    Exit Property
ErrorHandle:
    custErr.RaiseError "cHeaderRange - Set HeaderRange()"
End Property

Property Get HeaderRange() As Range
    On Error GoTo ErrorHandle
    If vHRng Is Nothing Then Err.Raise custErr.GenericError, Description:="The header range has not been defined"
    Set HeaderRange = vHRng
    Exit Property
ErrorHandle:
    custErr.RaiseError "cHeaderRange - Get HeaderRange()"
End Property


Property Get SheetRow() As Long
    On Error GoTo ErrorHandle
    SheetRow = Me.HeaderRange.row
    Exit Property
ErrorHandle:
    custErr.RaiseError "cHeaderRange - Get SheetRow()"
End Property


Property Get ColPresent(ByVal ColName As String) As Boolean
    On Error GoTo ErrorHandle
    ColPresent = Not Me.HeaderRange.Find(What:=ColName, LookIn:=XlFindLookIn.xlValues, LookAt:=XlLookAt.xlWhole, MatchCase:=True) Is Nothing
    Exit Property
ErrorHandle:
    custErr.RaiseError "cHeaderRange - Get ColPresent()"
End Property

Property Get ColCell(ByVal ColName As String) As Range
    On Error GoTo ErrorHandle
    If Not Me.ColPresent(ColName) Then Err.Raise custErr.GenericError, Description:="The given column name, " & ColName & ", could not be found in the header"
    Set ColCell = Me.HeaderRange.Find(What:=ColName, LookIn:=XlFindLookIn.xlValues, LookAt:=XlLookAt.xlWhole, MatchCase:=True)
    Exit Property
ErrorHandle:
    custErr.RaiseError "cHeaderRange - Get ColCell()"
End Property

Property Get ColInSheetNum(ByVal ColName As String) As Long
    On Error GoTo ErrorHandle
    ColInSheetNum = Me.ColCell.Column
    Exit Property
ErrorHandle:
    custErr.RaiseError "cHeaderRange - Get ColInSheetNum()"
End Property

Property Get ColInSheetLetter(ByVal ColName As String) As String
    On Error GoTo ErrorHandle
    ColInSheetLetter = Replace(Cells(1, Me.ColInSheetNum(ColName)).Address(RowAbsolute:=False, ColumnAbsolute:=False), 1, "")
    Exit Property
ErrorHandle:
    custErr.RaiseError "cHeaderRange - Get ColInSheetLetter()"
End Property

Property Get ColInData(ByVal ColName As String) As Long
    On Error GoTo ErrorHandle
    ColInData = Me.ColInSheetNum(ColName) - Me.HeaderRange.Columns(1).Column + 1
    Exit Property
ErrorHandle:
    custErr.RaiseError "cHeaderRange - Get ColInData()"
End Property

Property Get WholeSheetCol(ByValColName As String) As Range
    On Error GoTo ErrorHandle
    WholeSheetCol = Me.ColCell.EntireColumn
    Exit Property
ErrorHandle:
    custErr.RaiseError "cHeaderRange - Get WholeSheetCol()"
End Property
