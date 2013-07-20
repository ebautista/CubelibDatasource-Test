Module MADOMethod

    Public Const G_MAIN_PASSWORD = "wack2"


    Public Sub RstOpen(ByVal Source As String, _
                       ByRef ADOConnection As ADODB.Connection, _
                       ByRef RecordsetToOpen As ADODB.Recordset, _
                       ByVal CursorType As ADODB.CursorTypeEnum, _
                       ByVal LockType As ADODB.LockTypeEnum, _
              Optional ByVal CacheSize As Long = 1, _
              Optional ByVal MakeOffline As Boolean = False)

        Dim strTable

        strTable = Mid(UCase(Source), InStr(1, UCase(Source), "FROM") + 5)
        If InStr(1, strTable, " ") > 0 Then
            strTable = Mid(strTable, 1, InStr(1, strTable, " ") - 1)
        End If


        On Error GoTo Handler

        ' Close and set to nothing recordset
        If (RecordsetToOpen Is Nothing = False) Then
            If (RecordsetToOpen.State = ADODB.ObjectStateEnum.adStateOpen) Then
                RecordsetToOpen.Close()
            End If

            RecordsetToOpen = Nothing
        End If

        ' Create new instance of recordset
        RecordsetToOpen = New ADODB.Recordset

        ' Set recordset properties
        If (MakeOffline = True) Then
            RecordsetToOpen.CursorLocation = ADODB.CursorLocationEnum.adUseClient
        End If
        RecordsetToOpen.CacheSize = CacheSize


        ' Open recordset
        RecordsetToOpen.Open(Source, ADOConnection, CursorType, LockType)

        On Error GoTo 0


        If (MakeOffline = True) Then
            RecordsetToOpen.ActiveConnection = Nothing
        End If

        Exit Sub

Handler:
        Select Case Err.Number
            Case -2147467259    'Table ___ is exclusively locked by user 'Admin' on machine ____ .
                '            If MsgBox(Translate(2160) & " " & strTable & " " & Translate(2161) & _
                '                      vbCrLf & vbCrLf & Translate(2162), vbInformation + vbOKCancel) = vbOK Then
                '                Resume
                '            Else
                '            End If
            Case -2147217865
                '            If MsgBox(Translate(2160) & " " & strTable & " " & Translate(2163) & _
                '                     vbCrLf & vbCrLf & Translate(2162), vbInformation + vbOKCancel) = vbOK Then
                '                Resume
                '            Else
                '            End If
            Case Else
                ' Err.Raise Err.Number
        End Select
        On Error GoTo 0

    End Sub

    Public Sub RstClose(ByRef RecordsetToClose As ADODB.Recordset)

        If (RecordsetToClose Is Nothing = False) Then
            If (RecordsetToClose.State = ADODB.ObjectStateEnum.adStateOpen) Then
                RecordsetToClose.Close()
            End If

            RecordsetToClose = Nothing
        End If

    End Sub

    Public Sub DisconnectDB(ByRef ADOConnection As ADODB.Connection)

        If (ADOConnection Is Nothing = False) Then
            If ADOConnection.State = ADODB.ObjectStateEnum.adStateOpen Then
                ADOConnection.Close()
            End If

            ADOConnection = Nothing
        End If

    End Sub

    Public Sub ConnectDB(ByRef ADOConnection As ADODB.Connection, _
                         ByVal DBPath As String, _
                Optional ByVal DBName As String = vbNullString)

        If (ADOConnection Is Nothing = False) Then
            If (ADOConnection.State = ADODB.ObjectStateEnum.adStateOpen) Then
                ADOConnection.Close()
            End If

            ADOConnection = Nothing
        End If

        ADOConnection = New ADODB.Connection
        ADOConnection.Open("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & DBPath & IIf(DBName = vbNullString, vbNullString, "\" & DBName) & ";Persist Security Info=False;Jet OLEDB:Database Password=" & G_MAIN_PASSWORD)

    End Sub

End Module
