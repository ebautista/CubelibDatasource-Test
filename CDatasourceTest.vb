﻿Imports System.Text
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports CubelibDatasource
Imports ADODB.CursorTypeEnum
Imports ADODB.LockTypeEnum

<TestClass()> Public Class CDatasourceTest

    <TestMethod()> Public Sub TestUpdateSingleRecordQueried()
        Dim source As CDatasource = New CDatasource

        Dim conADO As ADODB.Connection = New ADODB.Connection
        Dim rstTemp As ADODB.Recordset = New ADODB.Recordset
        Dim rstTest As ADODB.Recordset = New ADODB.Recordset
        Dim success As Integer

        ConnectDB(conADO, My.Application.Info.DirectoryPath, "mdb_sadbel.mdb")
        Dim strCon As String = conADO.ConnectionString
        RstOpen("SELECT * FROM [PLDA IMPORT HEADER] WHERE [CODE] = '000000237661242485047'", conADO, rstTemp, adOpenKeyset, adLockOptimistic, , True)
        DisconnectDB(conADO)

        Assert.AreEqual(1, rstTemp.RecordCount)

        If rstTemp.RecordCount > 0 Then
            rstTemp.MoveFirst()

            'Store Original Values
            Dim strOriginalA4 As String = rstTemp.Fields("A4").Value
            Dim strOriginalA5 As String = rstTemp.Fields("A5").Value

            'Edit A4 and A5
            rstTemp.Fields("A4").Value = "20081206"
            rstTemp.Fields("A5").Value = "VP442020"

            'Use CubelibDatasource Update method
            Dim wrapperClass As New CRecordset(rstTemp, strCon)
            success = source.Update(wrapperClass, rstTemp.Bookmark, "PLDA IMPORT HEADER")
            Assert.IsTrue(success = 0)

            'Confirm that the changes has been saved to DB
            ConnectDB(conADO, My.Application.Info.DirectoryPath, "mdb_sadbel.mdb")
            RstOpen("SELECT * FROM [PLDA IMPORT HEADER] WHERE [CODE] = '000000237661242485047'", conADO, rstTest, adOpenKeyset, adLockOptimistic, , True)
            DisconnectDB(conADO)

            Assert.AreEqual(1, rstTest.RecordCount)
            Assert.AreEqual("20081206", rstTest.Fields("A4").Value)
            Assert.AreEqual("VP442020", rstTest.Fields("A5").Value)

            rstTest.MoveFirst()
            rstTest.Fields("A4").Value = strOriginalA4
            rstTest.Fields("A5").Value = strOriginalA5

            'Revert to original data
            wrapperClass = New CRecordset(rstTest, strCon)
            success = source.Update(wrapperClass, rstTest.Bookmark, "PLDA IMPORT HEADER")
            Assert.IsTrue(success = 0)

            RstClose(rstTemp)
            RstClose(rstTest)
            source = Nothing
        End If
    End Sub

    <TestMethod()> Public Sub TestUpdateMultipleRecordQueried()
        Dim source As CDatasource = New CDatasource

        Dim conADO As ADODB.Connection = New ADODB.Connection
        Dim rstTemp As ADODB.Recordset = New ADODB.Recordset
        Dim rstTest As ADODB.Recordset = New ADODB.Recordset
        Dim success As Integer

        ConnectDB(conADO, My.Application.Info.DirectoryPath, "mdb_sadbel.mdb")
        Dim strCon As String = conADO.ConnectionString
        RstOpen("SELECT * FROM [PLDA IMPORT HEADER]", conADO, rstTemp, adOpenKeyset, adLockOptimistic, , True)
        DisconnectDB(conADO)

        Assert.AreEqual(5, rstTemp.RecordCount)

        If rstTemp.RecordCount > 0 Then
            'Move recordset to third record
            rstTemp.MoveFirst()
            rstTemp.MoveNext()
            rstTemp.MoveNext()

            'Store Original Values
            Dim strOriginalA4 As String = rstTemp.Fields("A4").Value
            Dim strOriginalA5 As String = rstTemp.Fields("A5").Value
            Dim strCode As String = rstTemp.Fields("Code").Value

            'Edit A4 and A5
            rstTemp.Fields("A4").Value = "20081299"
            rstTemp.Fields("A5").Value = "VP442027"

            'Use CubelibDatasource Update method
            Dim wrapperClass As New CRecordset(rstTemp, strCon)
            success = source.Update(wrapperClass, rstTemp.Bookmark, "PLDA IMPORT HEADER")
            Assert.IsTrue(success = 0)

            'Confirm that the changes has been saved to DB
            ConnectDB(conADO, My.Application.Info.DirectoryPath, "mdb_sadbel.mdb")
            RstOpen("SELECT * FROM [PLDA IMPORT HEADER] WHERE [CODE] = '" & strCode & "'", conADO, rstTest, adOpenKeyset, adLockOptimistic, , True)
            DisconnectDB(conADO)

            Assert.AreEqual(1, rstTest.RecordCount)
            Assert.AreEqual("20081299", rstTest.Fields("A4").Value)
            Assert.AreEqual("VP442027", rstTest.Fields("A5").Value)

            rstTest.MoveFirst()
            rstTest.Fields("A4").Value = strOriginalA4
            rstTest.Fields("A5").Value = strOriginalA5

            'Revert to original data 
            wrapperClass = New CRecordset(rstTest, strCon)
            success = source.Update(wrapperClass, rstTest.Bookmark, "PLDA IMPORT HEADER")
            Assert.IsTrue(success = 0)

            RstClose(rstTemp)
            RstClose(rstTest)
            source = Nothing
        End If
    End Sub

    <TestMethod()> Public Sub TestUpdateMultipleRecordQueriedWithADOFilter()
        Dim source As CDatasource = New CDatasource

        Dim conADO As ADODB.Connection = New ADODB.Connection
        Dim rstTemp As ADODB.Recordset = New ADODB.Recordset
        Dim rstTest As ADODB.Recordset = New ADODB.Recordset
        Dim success As Integer

        ConnectDB(conADO, My.Application.Info.DirectoryPath, "mdb_sadbel.mdb")
        Dim strCon As String = conADO.ConnectionString
        RstOpen("SELECT * FROM [PLDA IMPORT HEADER]", conADO, rstTemp, adOpenKeyset, adLockOptimistic, , True)
        DisconnectDB(conADO)

        Assert.AreEqual(5, rstTemp.RecordCount)

        If rstTemp.RecordCount > 0 Then
            'filter record
            rstTemp.Filter = "Code = '000000958309650421142'"
            Assert.AreEqual(1, rstTemp.RecordCount)

            'Store Original Values
            Dim strOriginalA4 As String = rstTemp.Fields("A4").Value
            Dim strOriginalA5 As String = rstTemp.Fields("A5").Value
            Dim strCode As String = rstTemp.Fields("Code").Value

            'Edit A4 and A5
            rstTemp.Fields("A4").Value = "20081277"
            rstTemp.Fields("A5").Value = "VP442055"

            'Use CubelibDatasource Update method
            Dim wrapperClass As New CRecordset(rstTemp, strCon)
            success = source.Update(wrapperClass, rstTemp.Bookmark, "PLDA IMPORT HEADER")
            Assert.IsTrue(success = 0)

            'Confirm that the changes has been saved to DB
            ConnectDB(conADO, My.Application.Info.DirectoryPath, "mdb_sadbel.mdb")
            RstOpen("SELECT * FROM [PLDA IMPORT HEADER] WHERE [CODE] = '" & strCode & "'", conADO, rstTest, adOpenKeyset, adLockOptimistic, , True)
            DisconnectDB(conADO)
            Assert.AreEqual(1, rstTest.RecordCount)
            Assert.AreEqual("20081277", rstTest.Fields("A4").Value)
            Assert.AreEqual("VP442055", rstTest.Fields("A5").Value)

            rstTest.MoveFirst()
            rstTest.Fields("A4").Value = strOriginalA4
            rstTest.Fields("A5").Value = strOriginalA5

            'Revert to original data
            wrapperClass = New CRecordset(rstTest, strCon)
            success = source.Update(wrapperClass, rstTest.Bookmark, "PLDA IMPORT HEADER")
            Assert.IsTrue(success = 0)

            RstClose(rstTemp)
            RstClose(rstTest)
            source = Nothing
        End If
    End Sub

    <TestMethod()> Public Sub TestInsert()
        'Dim source As CDatasource = New CDatasource

        'Dim conADO As ADODB.Connection = New ADODB.Connection
        'Dim rstTemp As ADODB.Recordset = New ADODB.Recordset
        'Dim rstTest As ADODB.Recordset = New ADODB.Recordset
        'Dim success As Integer

        'ConnectDB(conADO, My.Application.Info.DirectoryPath, "mdb_sadbel.mdb")

        'RstOpen("SELECT * FROM [PLDA IMPORT HEADER]", conADO, rstTemp, adOpenKeyset, adLockOptimistic, , True)

        'Assert.AreEqual(5, rstTemp.RecordCount)

        'If rstTemp.RecordCount > 0 Then
        '    'Add new record to Recordset
        '    rstTemp.AddNew()
        '    'Assert.AreEqual(6, rstTemp.RecordCount)

        '    'Insert Values
        '    rstTemp.Fields("Code").Value() = "000000993949532508811"
        '    rstTemp.Fields("Header").Value() = "1"
        '    rstTemp.Fields("A1").Value() = "IM"
        '    rstTemp.Fields("A2").Value() = "Z"
        '    rstTemp.Fields("A3").Value() = "P945304849540810005111"
        '    rstTemp.Fields("A4").Value() = "20081288"
        '    rstTemp.Fields("A5").Value() = "VP442069"

        '    'Use CubelibDatasource Update method
        '    Dim wrapperClass As New CRecordset(rstTemp, conADO.ConnectionString)
        '    success = source.Insert(wrapperClass, rstTemp.Bookmark)
        '    Assert.IsTrue(success = 0)

        '    'Confirm that the changes has been saved to DB
        '    RstOpen("SELECT * FROM [PLDA IMPORT HEADER] WHERE [CODE] = '000000993949532508811'", conADO, rstTest, adOpenKeyset, adLockOptimistic)
        '    Assert.AreEqual(1, rstTest.RecordCount)
        '    Assert.AreEqual(1, rstTest.Fields("Header").Value)
        '    Assert.AreEqual("IM", rstTest.Fields("A1").Value)
        '    Assert.AreEqual("Z", rstTest.Fields("A2").Value)
        '    Assert.AreEqual("P945304849540810005111", rstTest.Fields("A3").Value)
        '    Assert.AreEqual("20081288", rstTest.Fields("A4").Value)
        '    Assert.AreEqual("VP442069", rstTest.Fields("A5").Value)

        '    'Remove inserted data 
        '    success = source.ExecuteNonQuery("DELETE FROM [PLDA IMPORT HEADER] WHERE [CODE] = '000000993949532508811'", CDatasource.DBInstanceType.DATABASE_SADBEL)
        '    Assert.IsTrue(success = 0)

        '    RstClose(rstTemp)
        '    RstClose(rstTest)
        '    source = Nothing
        'End If
    End Sub
End Class

