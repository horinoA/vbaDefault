VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DBcsv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Private mcon As ADODB.Connection
Private mrs As ADODB.Recordset
Private csvPath As String

Private Sub Class_Initialize()
    Set mcon = New ADODB.Connection
End Sub

Public Property Get getRs() As ADODB.Recordset
    Set getRs = mrs
End Property

Public Property Let setCsvPath(Path As String)
    csvPath = Path
    mcon.Open ("Provider=Microsoft.ACE.OLEDB.12.0;" & _
                 "Data Source=" & Path & ";" & _
                 "Extended Properties='Text;HDR=Yes;FMT=Delimited';")
End Property


Public Sub BeginTrans()
    mcon.BeginTrans
End Sub

Public Sub Commit()
    mcon.CommitTrans
End Sub

Public Sub Rollback()
    mcon.RollbackTrans
End Sub

Public Function rsSelect(strSQL As String) As Boolean

    On Error GoTo ERROR_rsSelect

    Set mrs = New ADODB.Recordset
    mrs.Open strSQL, mcon, adOpenKeyset, adLockOptimistic
    mrs.MoveFirst

    If Not mrs.BOF Then 'åüçıåãâ Ç™Ç†ÇÈ
        rsSelect = True
    Else
        rsSelect = False
    End If

    Exit Function

ERROR_rsSelect:

    rsSelect = False

End Function

Public Function rsExec(strSQL As String) As Boolean

    On Error GoTo ERROR_rsExec

    'Set mcon = CurrentProject.Connection
    Call mcon.Execute(strSQL)

    rsExec = True

    Exit Function

ERROR_rsExec:

    rsExec = False

End Function

Public Function rsOpen(strSQL As String) As Boolean

    On Error GoTo ERROR_rsOpen

    'Set mcon = CurrentProject.Connection
    Set mrs = New ADODB.Recordset
    mrs.Open strSQL, mcon, adOpenKeyset, adLockOptimistic

    rsOpen = True

    Exit Function

ERROR_rsOpen:

    rsOpen = False

End Function

Private Sub Class_Terminate()

    On Error Resume Next
    mcon.Close
    mrs.Close

    Set mcon = Nothing
    Set mrs = Nothing

End Sub

