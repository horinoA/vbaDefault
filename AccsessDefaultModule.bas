Attribute VB_Name = "AccsessDefaultModule"
Option Compare Database
'Microsoft ActiveX Data Objects *.*
'Microsoft Office *.* Object Library
'参照設定忘れんなよ('ω')ノ


Private Const HKEY_USERS As Long = &H80000003

Private Type GUID
 Data1 As Long
 Data2 As Integer
 Data3 As Integer
 Data4(7) As Byte
End Type

#If VBA7 And Win64 Then
  ' 64Bit 版
  Private Declare PtrSafe Function CoCreateGuid Lib "OLE32.DLL" (pGuid As GUID) As Long
  Private Declare PtrSafe Function QueryPerformanceCounter Lib "kernel32" _
                           (X As Currency) As Boolean
  Private Declare PtrSafe Function QueryPerformanceFrequency Lib "kernel32" _
                           (X As Currency) As Boolean

#Else
  ' 32Bit 版
  Private Declare Function CoCreateGuid Lib "OLE32.DLL" (pGuid As GUID) As Long
  Private Declare Function QueryPerformanceCounter Lib "Kernel32" _
                           (X As Currency) As Boolean
  Private Declare Function QueryPerformanceFrequency Lib "Kernel32" _
                           (X As Currency) As Boolean
#End If


Dim Freq As Currency
Dim Overhead  As Currency
Dim Ctr1 As Currency, Ctr2 As Currency, Result As Currency

'https://hatenachips.blog.fc2.com/blog-entry-377.html
'ミリ秒以下の高精度で処理時間計測
Public Sub SWStart()
    If QueryPerformanceCounter(Ctr1) Then
        QueryPerformanceCounter Ctr2
        QueryPerformanceFrequency Freq
        Overhead = Ctr2 - Ctr1
    Else
        Err.Raise 513, "StopwatchError", "High-resolution counter not supported."
    End If
    QueryPerformanceCounter Ctr1
End Sub

Public Sub SWStop()
    QueryPerformanceCounter Ctr2
    Result = (Ctr2 - Ctr1 - Overhead) / Freq * 1000
End Sub
 
Public Sub SWShow(Optional Caption As String)
    Debug.Print Caption & " " & Result
End Sub

'同じカレントパス内codeフォルダにVBAモジュール、クエリSQL,テーブル構造はく
Sub VBExport()

  Dim vbcmp As Object
  Dim sPath As String
  Dim strName As String
  Dim strExt As String
  Dim Qdf As DAO.QueryDef
  Dim dbs As DAO.Database
  Dim FNum As Integer
  Dim i As Integer
  
  sPath = CurrentProject.path & "¥code¥"
  
  For Each vbcmp In VBE.ActiveVBProject.VBComponents
    With vbcmp
      'ファイル名までを設定
      strName = sPath & .Name
      '拡張子を設定
      Select Case .Type
        Case 1    '標準モジュールの場合
          strExt = ".bas"
        Case 2    'クラスモジュールの場合
          strExt = ".cls"
        Case 100  'フォーム/レポートのモジュールの場合
          strExt = ".cls"
      End Select
      'モジュールをエクスポート
      .Export strName & strExt
    End With
  Next vbcmp
  
  'QuerySQL書き出し
  Set dbs = CurrentDb
  FNum = FreeFile
  strName = Mid(dbs.Name, InStrRev(dbs.Name, "¥") + 1)
  strName = Left(strName, InStrRev(strName, ".") - 1)
  strName = sPath & strName & "_Query.txt"
  Open strName For Output Access Write As #FNum
  
  For Each Qdf In dbs.QueryDefs

        'クエリ名＆SQLステートメント取得
        stSQL = "QueryName:" & Qdf.Name & vbCrLf & _
                "SQL:" & Qdf.SQL & vbCrLf & vbCrLf

        'ファイルに出力
        Print #FNum, stSQL

  Next
  Close #FNum
  
  'テーブル定義書き出し（システムファイル除く)
  FNum = FreeFile
  strName = Mid(dbs.Name, InStrRev(dbs.Name, "¥") + 1)
  strName = Left(strName, InStrRev(strName, ".") - 1)
  strName = sPath & strName & "_Table.txt"
  Open strName For Output Access Write As #FNum
  
  Dim myTD As TableDef
  
  For Each myTD In dbs.TableDefs
  
    If Left(myTD.Name, 2) <> "MS" Then
        stSQL = "TableName: " & myTD.Name & vbCrLf
        For Each myField In myTD.fields
            With myField
                stSQL = stSQL & "name: " & .Name & vbCrLf
                stSQL = stSQL & "Attributes: " & .Attributes & vbCrLf
                stSQL = stSQL & "CollatingOrder: " & .CollatingOrder & vbCrLf
                stSQL = stSQL & "Type: " & .Type & vbCrLf
                stSQL = stSQL & "OrdinalPosition: " & .OrdinalPosition & vbCrLf
                stSQL = stSQL & "Size: " & .Size & vbCrLf
                stSQL = stSQL & "SourceField: " & .SourceField & vbCrLf
                stSQL = stSQL & "SourceTable: " & .SourceTable & vbCrLf_
                stSQL = stSQL & "DataUpdatable: " & .DataUpdatable & vbCrLf
                stSQL = stSQL & "DefaultValue: " & .DefaultValue & vbCrLf
                stSQL = stSQL & "ValidationRule: " & .ValidationRule & vbCrLf
                stSQL = stSQL & "ValidationText: " & .ValidationText & vbCrLf
                stSQL = stSQL & "Required: " & .Required & vbCrLf
                stSQL = stSQL & "AllowZeroLength: " & .AllowZeroLength & vbCrLf & vbCrLf
            End With
        Next
        Print #FNum, stSQL
    End If

  Next
  
  Close #FNum
  
  Set dbs = Nothing

  MsgBox "コード吐き終わった"
  
End Sub



Public Function GetFileName(path As String) As String
'ファイルを開くダイアログの例

  Dim intRet As Integer
  With Application.FileDialog(msoFileDialogOpen)
    'ダイアログのタイトルを設定
    .title = "ファイルを開くダイアログの例"
    'ファイルの種類を設定
    .Filters.Clear
    .Filters.Add "テキストファイル", "*.txt;*.csv;*.prn"
    'ファイルの種類の初期値を設定
    .FilterIndex = 1
    '複数ファイル選択を許可しない
    .AllowMultiSelect = False
    '初期パスを設定
    .InitialFileName = path
    'ダイアログを表示
    intRet = .Show
    If intRet <> 0 Then
      'ファイルが選択されたとき
      'そのフルバスを返り値に設定
      GetFileName = Trim(.SelectedItems.Item(1))
    Else
      'ファイルが選択されなければ長さゼロの文字列を返す
      GetFileName = ""
    End If
  End With

End Function

'新規テキストファイル作成、uni=trueでunicode
Public Function fileCreate(FileName As String, Optional uni As Boolean) As Boolean
    
    On Error GoTo ERROR_fileCreate
    Dim flag As Boolean, fso As Object
    Debug.Print (FileName)
    If fileExist(FileName) = True Then
        flag = deleteFile(FileName)
        If flag = False Then    'ファイル削除エラー
            fileCreate = False
            Exit Function
        End If
    End If
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    If IsMissing(uniCode) = True Then
        Set ts = fso.CreateTextFile(FileName, Overwrite:=True, uniCode:=True)     'UNICODE
    Else
        Set ts = fso.CreateTextFile(FileName, Overwrite:=True, uniCode:=False)     ' other
    End If
    ts.Close
    
    Set ts = Nothing
    Set fso = Nothing
    fileCreate = True
    Exit Function
    
ERROR_fileCreate:

    ts.Close
    Set ts = Nothing
    Set fso = Nothing
    fileCreate = False
    Call Err.Raise(50000, "fileCreate", "ファイル作成エラー")
    
End Function

'ファイル(path)が存在するか
Public Function fileExist(path As String) As Boolean
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    fileExist = fso.FileExists(path)
    Set fso = Nothing
End Function

'指定文字strを指定ファイルfileNameに書き込む
Public Sub fileWrite(str As String, fileName As String)
        
    On Error GoTo ERR_fileWrite
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    With fso.getFile(fileName).OpenAsTextStream(8)
        .WriteLine str
        .Close
    End With
    Set fso = Nothing
    Exit Sub
    
ERR_fileWrite:
    Set fso = Nothing
    Call Err.Raise(60000, "fileWrite", "ファイル書き込みエラー")
        
End Sub

'テーブルが存在するか
Public Function ExistTable(tableName As String) As Boolean
    On Error Resume Next
    Dim myDB As DAO.Database, myTD As DAO.TableDef, flag As Boolean
    Set myDB = CurrentDb
    
    For Each myTD In myDB.TableDefs
        If myTD.Name = tableName Then
            ExistTable = True
            Set myDB = Nothing
            Exit Function
        End If
    Next
    ExistTable = False
    Set myDB = Nothing
End Function

'クエリが存在するか
Public Function ExistQuery(queryName As String) As Boolean
    Dim Obj As AccessObject
    For Each Obj In CurrentData.AllQueries
        If Obj.Name = queryName Then
            ExistQuery = True
            Exit Function
        End If
    Next
    ExistQuery = False
End Function

'指定したフォルダ(path)よりファイル名(fileName)が含まれるファイルで一番直近のファイルパスを返す
Public Function fileNameMaxday(fileName As String, path As String) As String
    Dim fso, fl, f
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set fl = fso.GetFolder(path)
    Dim fileName1 As String, d As Date

    For Each f In fl.Files ' フォルダ内のファイルを取得
        If f.Name Like "*" & fileName & "*" And f.DateCreated > d Then  ' 日時を取得したいファイル
            d = f.DateCreated      ' 作成日時を取得
            fileName1 = f.Name
        End If
    Next

    Set fso = Nothing
    
    fileNameMaxday = CStr(fileName1)

End Function

'CSV(csvName),WHERE条件(strWhere,""空白文字なら条件なし)より指定したテーブル(tabkeName)にデータインサート。dbにより外部Accsess指定可
Sub CSVtoTable(csvName As String, strWhere As String, tableName As String, Optional db As Variant)

    On Error GoTo ERR_CSVtoTable
    
    Dim fields As Variant, rsFlag As Boolean, i As Long, j As Long
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim rs As DAO.RecordSet, tdf As TableDef, fld As Field
    
    If IsMissing(db) Then
        Set rs = CurrentDb.OpenRecordset(tableName)
        Set tdf = CurrentDb.TableDefs(tableName)
    Else
        Set rs = db.OpenRecordset(tableName)
        Set tdf = db.TableDefs(tableName)
    End If
    
    Dim con As DBcsv
    Set con = New DBcsv
    
    con.setCsvPath = fso.GetParentFolderName(csvName) & "¥"
    
    Dim strSQL As String
    
    strSQL = "SELECT * FROM [" & fso.GetFileName(csvName) & "]"
    If strWhere <> "" Then
        strSQL = strSQL & " WHERE " & strWhere
    End If
    rsFlag = con.rsSelect(strSQL)
    
    If rsFlag = True Then
        Do While Not con.getRs.EOF
            rs.AddNew
            For Each conFld In con.getRs.fields
                fldName = conFld.Name
                
                
                '追加先レコードセットにフィールド名があるか？
                For i = 0 To rs.fields.Count - 1
                    If rs.fields(i).Name = fldName Then
                        '日付型
                        If rs.fields(i).Type = dbDate Or rs.fields(i).Type = dbTime Then
                            rs.fields(fldName) = cstAccsessTime(CStr(conFld.Value))
                        'テキスト
                        ElseIf rs.fields(i).Type = dbChar Or rs.fields(i).Type = dbDouble Or rs.fields(i).Type = dbMemo Or _
                            rs.fields(i).Type = dbText Or rs.fields(i).Type = dbSingle Then
                            rs.fields(fldName) = Nz(conFld.Value, "")
                        '数値
                        Else
                            rs.fields(fldName) = Nz(conFld.Value, 0)
                        End If
                        Exit For
                    End If
                Next i
                
            Next conFld
            
            rs.Update
            con.getRs.MoveNext
        Loop
    End If
    
    Set fso = Nothing
    Set con = Nothing
    rs.Close: Set rs = Nothing
    Exit Sub
    
ERR_CSVtoTable:

    Set fso = Nothing
    Call Err.Raise(60001, "CSVtoTable", "CSVtoTable変換エラー")
    
End Sub

'ADOを使ってSQL文より（tableCreateSQL）新規テーブル作成、10進数型などフィールドに指定できる
Function tableInit(tableName As String, tableCreateSQL As String)

    Dim tablecheck As Boolean
    
    tablecheck = ExistTable(tableName)
    If tablecheck = True Then
        DoCmd.RunSQL "DELETE FROM " & tableName & ";"
    Else
        'ADOにてCREATE
        Dim cn As New ADODB.Connection
        Set cn = CurrentProject.Connection
        Dim cm As New ADODB.Command
    
        cm.ActiveConnection = cn
        cm.CommandText = tableCreateSQL
        cm.Execute

        Set cm = Nothing
        cn.Close: Set cn = Nothing
    End If

End Function

'PostgressqlSQL文（strSQL）より,登録先テーブル(tableName)に追加、用DBCore
Function getPostgresToTable(strSQL As String, tableName As String)
    
    On Error GoTo PostgresToTable
    
    Dim rs As DAO.RecordSet
    Dim conFld As ADODB.Field
    Dim rsFlag As Boolean 'レコード判定
    Dim con As DBCore
    Set con = New DBCore
    Dim fldName As String, i As Long
    

    Set rs = CurrentDb.OpenRecordset(tableName)
    Set tbf = CurrentDb.TableDefs(tableName)
    rsFlag = con.rsSelect(strSQL)

    If rsFlag = True Then
        con.getRs.MoveFirst
       
        Do While Not con.getRs.EOF
            rs.AddNew
            For Each conFld In con.getRs.fields
                fldName = conFld.Name
                '追加先レコードセットにフィールド名があるか？
                For i = 0 To rs.fields.Count - 1
                    If rs.fields(i).Name = fldName Then
                        rs.fields(fldName) = conFld.Value
                        Exit For
                    Else
                    End If
                Next i
            Next conFld
            rs.Update
            con.getRs.MoveNext
        Loop
        
    Else
        Set rs = Nothing
        Set tbf = Nothing
        Set con = Nothing
        Err.Raise Number:=1021, Description:="SQL該当レコードがありませんでした。"
        Exit Function
    End If
    
    Set rs = Nothing
    Set tbf = Nothing
    Set con = Nothing
    
    Exit Function
    
PostgresToTable:
    
    Set rs = Nothing
    Set tbf = Nothing
    Set con = Nothing
    
    Dim lsMsg As String, logErrorFile As String
    Dim DbErr As Error
    lsMsg = ""
    If DBEngine.Errors.Count > 0 Then
        For Each DbErr In DBEngine.Errors
            lsMsg = lsMsg & DbErr.Description & vbCrLf
        Next DbErr
    End If
    If lsMsg = "" Then lsMsg = "postgres接続中エラーが発生しました。（SQLエラー又はタイムアウト)"
    
    Err.Raise Number:=1020, Description:=lsMsg
    
End Function

'外部Accsess（dbName）,SQL文(strSQL)、より登録先テーブル（tableName）に追加
Function getAcsessToTable(strSQL As String, tableName As String, dbName As String)
    
    On Error GoTo AcsessToTable
    
    Dim rs As DAO.RecordSet
    Dim conFld As ADODB.Field
    Dim rsFlag As Boolean 'レコード判定
    Dim con As DBAccss
    Set con = New DBAccss
    con.setAccsessFile = dbName
    Dim fldName As String, i As Long
    

    Set rs = CurrentDb.OpenRecordset(tableName)
    Set tbf = CurrentDb.TableDefs(tableName)
    rsFlag = con.rsSelect(strSQL)

    If rsFlag = True Then
        con.getRs.MoveFirst
       
        Do While Not con.getRs.EOF
            rs.AddNew
            For Each conFld In con.getRs.fields
                fldName = conFld.Name
                '追加先レコードセットにフィールド名があるか？
                For i = 0 To rs.fields.Count - 1
                    If rs.fields(i).Name = fldName Then
                        rs.fields(fldName) = conFld.Value
                        Exit For
                    End If
                Next i
            Next conFld
            rs.Update
            con.getRs.MoveNext
        Loop
        
    Else
        Set rs = Nothing
        Set tbf = Nothing
        Set con = Nothing
        Err.Raise Number:=1021, Description:="SQL該当レコードがありませんでした。"
        Exit Function
    End If
    
    Set rs = Nothing
    Set tbf = Nothing
    Set con = Nothing
    
    Exit Function
    
AcsessToTable:
    
    Set rs = Nothing
    Set tbf = Nothing
    Set con = Nothing
    
    Dim DbErr As Error, str As String
    
    If DBEngine.Errors.Count > 0 Then
        For Each DbErr In DBEngine.Errors
            If DbErr.Number <> 3146 Then
                str = str & DbErr.Description & vbCrLf
            End If
        Next DbErr
    End If
    
    Err.Raise Number:=1020, Description:="Acsess接続に失敗しました" & vbCrLf & str
End Function


'フルパス(FileName)よりファイル名のみ取り出す
Function GetFileNameFromPath(FileName As String) As String
    With CreateObject("Scripting.FileSystemObject")
        Dim strNum As Long
        strNum = Len(.getFile(FileName).Name) - Len(.GetExtensionName(FileName)) - 1
        Debug.Print (Left(.getFile(FileName).Name, strNum))
        GetFileNameFromPath = Left(.getFile(FileName).Name, strNum)
    End With
End Function


'GUID生成
'https://kazuyaujihara.wordpress.com/2017/08/27/vba-%E3%81%A7-guid-%E3%82%92%E7%94%9F%E6%88%90%E3%81%99%E3%82%8B/
Public Function getGUID() As String
 '(c) 2000 Gus Molina
 Dim udtGUID As GUID
 If (CoCreateGuid(udtGUID) = 0) Then
  getGUID = _
   String(8 - Len(Hex$(udtGUID.Data1)), "0") & Hex$(udtGUID.Data1) & "-" & _
   String(4 - Len(Hex$(udtGUID.Data2)), "0") & Hex$(udtGUID.Data2) & "-" & _
   String(4 - Len(Hex$(udtGUID.Data3)), "0") & Hex$(udtGUID.Data3) & "-" & _
   IIf((udtGUID.Data4(0) < &H10), "0", "") & Hex$(udtGUID.Data4(0)) & _
   IIf((udtGUID.Data4(1) < &H10), "0", "") & Hex$(udtGUID.Data4(1)) & "-" & _
   IIf((udtGUID.Data4(2) < &H10), "0", "") & Hex$(udtGUID.Data4(2)) & _
   IIf((udtGUID.Data4(3) < &H10), "0", "") & Hex$(udtGUID.Data4(3)) & _
   IIf((udtGUID.Data4(4) < &H10), "0", "") & Hex$(udtGUID.Data4(4)) & _
   IIf((udtGUID.Data4(5) < &H10), "0", "") & Hex$(udtGUID.Data4(5)) & _
   IIf((udtGUID.Data4(6) < &H10), "0", "") & Hex$(udtGUID.Data4(6)) & _
   IIf((udtGUID.Data4(7) < &H10), "0", "") & Hex$(udtGUID.Data4(7))
 End If
End Function

'指定したクエリを削除qName
Public Function DeleteView(qName As String) As Boolean
    Dim cn As New ADODB.Connection
    Dim cat As New ADOX.Catalog
    Dim vew As ADOX.View
    Dim DataFlag As Integer, flag As Boolean
 
    Set cn = CurrentProject.Connection
    Set cat.ActiveConnection = cn
    'クエリの存在チェック
    For Each vew In cat.Views
       If vew.Name = qName Then
         DataFlag = 1
       End If
    Next vew
    'クエリが存在している場合は削除
    If DataFlag = 1 Then
      cat.Views.Delete (qName)
      GoTo END_DeleteView
    Else
      GoTo END_DeleteView
    End If
    
END_DeleteView:
    cn.Close
    Set cn = Nothing
    Set cat = Nothing
    If DataFlag = 1 Then
        DeleteView = True
    Else
        DeleteView = False
    End If
End Function


'レコードセットからCSVファイル作成
Function outPutCSV(rs As DAO.RecordSet, path As String)
    
    On Error GoTo ERROR_CSV
    
    Dim i As Long, fieldType As Integer, lineStr As String, fieldsName As String, str As String, num As Variant
    Dim lngFileNum As Long
    
    lngFileNum = FreeFile()
    Open path For Output As #lngFileNum
    
    For i = 0 To rs.fields.Count - 1
        fieldsName = rs.fields(i).Name
        If InStr(fieldsName, ".") > 0 Then
            fieldsName = Replace(fieldsName, ".", "_")
        End If
        lineStr = lineStr & fieldsName
        If i <> rs.fields.Count - 1 Then
            lineStr = lineStr & ","
        End If
    Next i
    Print #lngFileNum, lineStr

    rs.MoveFirst
    Do While Not rs.EOF
        lineStr = ""
               
        For i = 0 To rs.fields.Count - 1
            fieldType = rs.fields(i).Type
            If fieldType = dbText Or fieldType = dbMemo Or fieldType = dbChar Then
                str = IIf(IsNull(rs.fields(i)), "", rs.fields(i))
                lineStr = lineStr & """" & str & """"
            Else
                num = IIf(IsNull(rs.fields(i)), 0, rs.fields(i))
                lineStr = lineStr & CStr(num)
            End If
            If i <> rs.fields.Count - 1 Then
                lineStr = lineStr & ","
            End If
        Next i
        
        Print #lngFileNum, lineStr
        rs.MoveNext
    Loop
    
    Close #lngFileNum
    Exit Function
    
ERROR_CSV:
    
    Close #lngFileNum
    Call Err.Raise(70000, "outPutCSV", "RecordSettoCSV作成エラー")

End Function

'pathで指定したファイルを削除する
Function deleteFile(path As String) As Boolean
    On Error GoTo ERROR_deleteFile
    Dim flag As Boolean, fso As Object
    flag = fileExist(path)
    If flag = True Then
        Set fso = CreateObject("Scripting.FileSystemObject")
        Call fso.deleteFile(path, True)
        Set fso = Nothing
        deleteFile = True
    Else
        deleteFile = False
    End If
    Exit Function
    
ERROR_deleteFile:
    deleteFile = False
End Function

'正規表現で開始位置チェック
Function regMatchIndex(pat As String, str As String) As Long
    On Error GoTo regFaild
    
    If pat = "" Or str = "" Then
        GoTo regFaild
    End If
    
    Dim reg, matches, setStr As String
    Set reg = CreateObject("VBScript.RegExp")
    With reg
        .pattern = pat
        .IgnoreCase = True
        .Global = True
    End With
    Set matches = reg.Execute(str)
    regMatchIndex = matches.Item(0).firstindex
    Set reg = Nothing
    Set matches = Nothing
    Exit Function
    
regFaild:
    regMatchIndex = -1
    Set reg = Nothing
    Set matches = Nothing
End Function

'テキストから指定した単位についた数字部を取り出す、正規表現
Function getNumfromText(str As String, unit As String) As Double
    On Error GoTo getNumERROR
    Dim i As Long, txt As String
    i = regMatchIndex(unit, str)
    Do
        If Mid(str, i, 1) Like "[0-9]" Or Mid(str, i, 1) Like "[.]" Then
            txt = Mid(str, i, 1) & txt
        Else
            Exit Do
        End If
        i = i - 1
    Loop While i > 0
    getNumfromText = CDbl(txt)
    Exit Function

getNumERROR:
    getNumfromText = 0#
End Function

'文字列が含まれたらその文字列を返す、正規表現
Function getTxtfromText(str As String, unit As String) As String
    On Error GoTo getTxtERROR
    Dim i As Long, txt As String
    i = regMatchIndex(unit, str)
    If i >= 0 Then
        getTxtfromText = unit
    Else
        getTxtfromText = ""
    End If
    Exit Function
getTxtERROR:
    getTxtfromText = ""
End Function

'レコードセットから配列作成用文字列返す
Public Function rsToArrry(rs As ADODB.RecordSet, fieldName As String) As String
    Dim productArry As String
    rs.MoveFirst
    Do While Not rs.EOF
        productArry = productArry & CStr(rs.fields(fieldName)) & ","
        rs.MoveNext
    Loop
    productArry = Left(productArry, Len(productArry) - 1)
    rsToArrry = productArry
End Function


'PostgresからAccsessに変換する際何もしないと日付の丸めが起こるためそれを防ぐSQLに変換
'to_timestamp(to_char(フィールド名, 'yyyy/mm/dd hh24:mi:ss'), 'yyyy/mm/dd hh24:mi:ss')
Public Function acsesDayCastSQL(fieldName As String) As String
    If IsNull(fieldName) Or fieldName = "" Then
        acsesDayCastSQL = ""
        Exit Function
    End If
    
    acsesDayCastSQL = "to_timestamp(to_char(" & fieldName & ", 'yyyy/mm/dd hh24:mi:ss'), 'yyyy/mm/dd hh24:mi:ss')"
End Function

'該当テーブル単体(tableNmae)のリンクをリフレッシュ、pathで指定したフルパスでリンクを付け替える
Function refreshLinkTable(tableName As String, Optional path As String = "") As Boolean
    
    On Error GoTo ERROR_refreshLink
    Dim tbl As DAO.TableDef, db As DAO.Database
    
    Set db = CurrentDb()
    Set tbl = db.TableDefs(tableName)
    
    If path <> "" Then
        tbl.Connect = ";DATABASE=" & path
    End If
    
    tbl.RefreshLink
    Set db = Nothing
    Set tbl = Nothing
    
    refreshLinkTable = True
    Exit Function
    
ERROR_refreshLink:
    Set db = Nothing
    Set tbl = Nothing
    refreshLinkTable = False
End Function

'新規リンクテーブル作成
Function appendLinkTable(tb As String, db As String) As Boolean
    
    On Error GoTo ERROR_appendLink
    
    If ExistTable(tb) = True Then
        appendLinkTable = False
        Exit Function
    End If
    
    Dim tbl As DAO.TableDef
    Set tbl = CurrentDb.CreateTableDef(tb)
    tbl.Connect = ";DATABASE=" & db
    tbl.SourceTableName = tb 'リンク元テーブル名
    CurrentDb.TableDefs.Append tbl
    
    Set tbl = Nothing
    appendLinkTable = True
    Exit Function

ERROR_appendLink:
    appendLinkTable = False
End Function

'全リンクテーブルリンク再接続
Function linkTableRefresh()

    Dim dbs As DAO.Database
    Dim tdf As DAO.TableDef
    
    Set dbs = CurrentDb
    
    '全テーブルの探索ループ
    For Each tdf In dbs.TableDefs
    With tdf
      If .Attributes And dbAttachedTable Then
        .RefreshLink
      End If
    End With
    Next tdf

End Function

'Accsessでなんちゃってoracleのconcat関数
'http://allenbrowne.com/func-concat.html
Public Function ConcatRelated(strField As String, _
    strTable As String, _
    Optional strWhere As String, _
    Optional strOrderBy As String, _
    Optional strSeparator = "_") As Variant
On Error GoTo Err_Handler
    'Purpose:   Generate a concatenated string of related records.
    'Return:    String variant, or Null if no matches.
    'Arguments: strField = name of field to get results from and concatenate.
    '           strTable = name of a table or query.
    '           strWhere = WHERE clause to choose the right values.
    '           strOrderBy = ORDER BY clause, for sorting the values.
    '           strSeparator = characters to use between the concatenated values.
    'Notes:     1. Use square brackets around field/table names with spaces or odd characters.
    '           2. strField can be a Multi-valued field (A2007 and later), but strOrderBy cannot.
    '           3. Nulls are omitted, zero-length strings (ZLSs) are returned as ZLSs.
    '           4. Returning more than 255 characters to a recordset triggers this Access bug:
    '               http://allenbrowne.com/bug-16.html
    Dim rs As DAO.RecordSet         'Related records
    Dim rsMV As DAO.RecordSet       'Multi-valued field recordset
    Dim strSQL As String            'SQL statement
    Dim strOut As String            'Output string to concatenate to.
    Dim lngLen As Long              'Length of string.
    Dim bIsMultiValue As Boolean    'Flag if strField is a multi-valued field.
    
    'Initialize to Null
    ConcatRelated = Null
    
    'Build SQL string, and get the records.
    strSQL = "SELECT " & strField & " FROM " & strTable
    If strWhere <> vbNullString Then
        strSQL = strSQL & " WHERE " & strWhere
    End If
    If strOrderBy <> vbNullString Then
        strSQL = strSQL & " ORDER BY " & strOrderBy
    End If
    Set rs = DBEngine(0)(0).OpenRecordset(strSQL, dbOpenDynaset)
    'Determine if the requested field is multi-valued (Type is above 100.)
    bIsMultiValue = (rs(0).Type > 100)
    
    'Loop through the matching records
    Do While Not rs.EOF
        If bIsMultiValue Then
            'For multi-valued field, loop through the values
            Set rsMV = rs(0).Value
            Do While Not rsMV.EOF
                If Not IsNull(rsMV(0)) Then
                    strOut = strOut & rsMV(0) & strSeparator
                End If
                rsMV.MoveNext
            Loop
            Set rsMV = Nothing
        ElseIf Not IsNull(rs(0)) Then
            strOut = strOut & rs(0) & strSeparator
        End If
        rs.MoveNext
    Loop
    rs.Close
    
    'Return the string without the trailing separator.
    lngLen = Len(strOut) - Len(strSeparator)
    If lngLen > 0 Then
        ConcatRelated = Left(strOut, lngLen)
    End If

Exit_Handler:
    'Clean up
    Set rsMV = Nothing
    Set rs = Nothing
    Exit Function

Err_Handler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbExclamation, "ConcatRelated()"
    Resume Exit_Handler
End Function


'文字列をAccsessのフィールドに入る日付型yyyy/mm/dd hhmmss形式に変換
Function cstAccsessTime(str As String) As Date
    On Error GoTo ERRORcstAccsessTime
    
    If str = "" Then
        GoTo ERRORcstAccsessTime
    End If
    
    Dim reg, matches, m, setStr As String
    Set reg = CreateObject("VBScript.RegExp")
    
    With reg
        .pattern = "(¥d{4})/(¥d{2})/(¥d{2}) (¥d{2}):(¥d{2}):(¥d{2})"    '日付＆時刻
        .IgnoreCase = True
        .Global = True
    End With
    
    If reg.test(str) = False Then
        reg.pattern = "(¥d{4})/(¥d{2})/(¥d{2})" '日付
        If reg.test(str) = False Then
            reg.pattern = "(¥d{2}):(¥d{2}):(¥d{2})" '時刻
            Set matches = reg.Execute(str)
            For Each m In matches
                setStr = m.Value
            Next
        Else
            Set matches = reg.Execute(str)
            For Each m In matches
                setStr = m.Value
            Next
        End If
    Else
        Set matches = reg.Execute(str)
        For Each m In matches
            setStr = m.Value
        Next
    End If
        
    cstAccsessTime = CDate(setStr)
    
    Set reg = Nothing
    Set matches = Nothing
    Exit Function
    
ERRORcstAccsessTime:
    cstAccsessTime = 0
    Set reg = Nothing
    Set matches = Nothing
End Function


'新規アクセスファイル作成（filePath フルパス）
Function makeAccsessDB(filePath As String) As Boolean
    
    On Error GoTo ERRORmakeAccsessDB
    
    With New Application
        .NewCurrentDatabase filePath
        .Quit
    End With
    makeAccsessDB = True
    Exit Function
    
ERRORmakeAccsessDB:
    makeAccsessDB = False
End Function

'https://vbabeginner.net/convert-utf-8-files-to-shift-jis/
'文字コードutf->shift-jis
Function Utf8ToSjis(a_sFrom, a_sTo) As Boolean

    On Error GoTo ERRORUtf8ToSjis
    
    Dim streamRead  As New ADODB.Stream '// 読み込みデータ
    Dim streamWrite As New ADODB.Stream '// 書き込みデータ
    Dim sText                           '// ファイルデータ
    
    streamRead.Type = adTypeText
    streamRead.Charset = "UTF-8"
    streamRead.Open
    Call streamRead.LoadFromFile(a_sFrom)
    
    '// 改行コードLFをCRLFに変換
    sText = streamRead.ReadText
    sText = Replace(sText, vbLf, vbCrLf)
    sText = Replace(sText, vbCr & vbCr, vbCr)
    
    streamWrite.Type = adTypeText
    streamWrite.Charset = "Shift-JIS"
    streamWrite.Open
    
    Call streamWrite.WriteText(sText)
    Call streamWrite.SaveToFile(a_sTo, adSaveCreateOverWrite)
    
    streamRead.Close
    streamWrite.Close
    
    Utf8ToSjis = True
    Exit Function
    
ERRORUtf8ToSjis:
    
    Utf8ToSjis = False
    
End Function

'数字8桁を日付変換
Function conv8digTodate(str As String) As Date

    Dim reg
    Set reg = CreateObject("VBScript.RegExp")

    With reg
        .pattern = "¥d{8}"
        .IgnoreCase = True
        .Global = True
    End With
    
    If reg.test(str) = False Then
        conv8digTodate = 0
        Set reg = Nothing
        Exit Function
    End If

    conv8digTodate = DateSerial(CInt(Left(str, 4)), CInt(Mid(str, 5, 2)), CInt(Mid(str, 7, 2)))

End Function

'https://popozure.info/20190515/14201
'fncGetCharset Ver1.6 @popozure
'ファイルの文字コード判定を行う
Function fncGetCharset(FileName As String) As String
    Dim i                   As Long     '汎用指数
       
    Dim lngFileLen          As Long     'ファイルサイズ
    Dim bytFile()           As Byte     'ファイル内容
    Dim b1                  As Byte     '1バイト目
    Dim b2                  As Byte     '2バイト目
    Dim b3                  As Byte     '3バイト目
    Dim b4                  As Byte     '4バイト目
       
    Dim lngSJIS             As Long     'Shift_JISの可能性
    Dim lngUTF8             As Long     'UTF-8もの可能性
    Dim lngEUC              As Long     'EUC-JPの可能性
     
    'ADODB定数
    Const adModeUnknown = 0
    Const adModeRead = 1
    Const adModeWrite = 2
    Const adModeReadWrite = 3
    Const adModeShareDenyRead = 4
    Const adModeShareDenyWrite = 8
    Const adModeShareExclusive = 12
    Const adModeShareDenyNone = 16
    Const adTypeBinary = 1
    Const adTypeText = 2
    Const adReadAll = -1
    Const adReadLine = -2
     
    'ファイル読み込み（バイナリー）
    On Error Resume Next
    With CreateObject("ADODB.Stream")
        .Mode = adModeUnknown
        .Open
        .Type = adTypeBinary
        .LoadFromFile FileName
        lngFileLen = .Size
        bytFile = .Read(adReadAll)
        .Close
    End With
    If (Err.Number <> 0) Then
        fncGetCharset = "OPEN FAILED"
        Exit Function
    End If
    On Error GoTo 0
     
    'BOMによる判断
    If (bytFile(0) = &HEF And bytFile(1) = &HBB And bytFile(2) = &HBF) Then
        fncGetCharset = "UTF-8 BOM"
        Exit Function
    ElseIf (bytFile(0) = &HFF And bytFile(1) = &HFE) Then
        fncGetCharset = "UTF-16 LE BOM"
        Exit Function
    ElseIf (bytFile(0) = &HFE And bytFile(1) = &HFF) Then
        fncGetCharset = "UTF-16 BE BOM"
        Exit Function
    End If
       
    'BINARY
    For i = 0 To lngFileLen - 1
        b1 = bytFile(i)
        If ((b1 >= &H0 And b1 <= &H1F) And b1 <> &H9 And b1 <> &HA And b1 <> &HD And b1 <> &H1B) Or (b1 = &H7F) Then
            fncGetCharset = "BINARY"
            Exit Function
        End If
    Next i
              
    'SJIS
    For i = 0 To lngFileLen - 1
        b1 = bytFile(i)
        If (b1 = &H9) Or (b1 = &HA) Or (b1 = &HD) Or (b1 >= &H20 And b1 <= &H7E) Or (b1 >= &HB0 And b1 <= &HDF) Then
            lngSJIS = lngSJIS + 1
        Else
            If (i < lngFileLen - 2) Then
                b2 = bytFile(i + 1)
                If ((b1 >= &H81 And b1 <= &H9F) Or (b1 >= &HE0 And b1 <= &HFC)) And _
                   ((b2 >= &H40 And b2 <= &H7E) Or (b2 >= &H80 And b2 <= &HFC)) Then
                   lngSJIS = lngSJIS + 2
                   i = i + 1
                End If
            End If
        End If
    Next i
              
    'UTF-8
    For i = 0 To lngFileLen - 1
        b1 = bytFile(i)
        If (b1 = &H9) Or (b1 = &HA) Or (b1 = &HD) Or (b1 >= &H20 And b1 <= &H7E) Then
            lngUTF8 = lngUTF8 + 1
        Else
            If (i < lngFileLen - 2) Then
                b2 = bytFile(i + 1)
                If (b1 >= &HC2 And b1 <= &HDF) And (b2 >= &H80 And b2 <= &HBF) Then
                   lngUTF8 = lngUTF8 + 2
                   i = i + 1
                Else
                    If (i < lngFileLen - 3) Then
                        b3 = bytFile(i + 2)
                        If (b1 >= &HE0 And b1 <= &HEF) And (b2 >= &H80 And b2 <= &HBF) And (b3 >= &H80 And b3 <= &HBF) Then
                            lngUTF8 = lngUTF8 + 3
                            i = i + 2
                        Else
                            If (i < lngFileLen - 4) Then
                                b4 = bytFile(i + 3)
                                If (b1 >= &HF0 And b1 <= &HF7) And (b2 >= &H80 And b2 <= &HBF) And (b3 >= &H80 And b3 <= &HBF) And (b4 >= &H80 And b4 <= &HBF) Then
                                    lngUTF8 = lngUTF8 + 4
                                    i = i + 3
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    Next i
   
    'EUC-JP
    For i = 0 To lngFileLen - 1
        b1 = bytFile(i)
        If (b1 = &H9) Or (b1 = &HA) Or (b1 = &HD) Or (b1 >= &H20 And b1 <= &H7E) Then
            lngEUC = lngEUC + 1
        Else
            If (i < lngFileLen - 2) Then
                b2 = bytFile(i + 1)
                If ((b1 >= &HA1 And b1 <= &HFE) And _
                   (b2 >= &HA1 And b2 <= &HFE)) Or _
                   ((b1 = &H8E) And (b2 >= &HA1 And b2 <= &HDF)) Then
                   lngEUC = lngEUC + 2
                   i = i + 1
                End If
            End If
        End If
    Next i
              
    '文字コード出現順位による判断
    If (lngSJIS <= lngUTF8) And (lngEUC <= lngUTF8) Then
        fncGetCharset = "UTF-8"
        Exit Function
    End If
    If (lngUTF8 <= lngSJIS) And (lngEUC <= lngSJIS) Then
        fncGetCharset = "Shift_JIS"
        Exit Function
    End If
    If (lngUTF8 <= lngEUC) And (lngSJIS <= lngEUC) Then
        fncGetCharset = "EUC-JP"
        Exit Function
    End If
     
    '判定不能
    fncGetCharset = "UNKNOWN"
End Function

