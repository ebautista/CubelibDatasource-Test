Imports System.Text
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports CubelibDatasource
Imports ADODB.CursorTypeEnum
Imports ADODB.LockTypeEnum
Imports CubelibDatasource.CDatasource

<TestClass()> Public Class CDatasourceTest

    Private source As CDatasource = New CDatasource

    <TestInitialize()> Public Sub Init()
        source.ExecuteNonQuery("INSERT INTO [PLDA IMPORT HEADER] ([Code], [Header], [A1], [A2], [A3], [A4], [A5]) VALUES  ('000000993949532508849', 1, 'IM', 'Z', 'P945304849540810005246', '20081206', 'VP442020')", CDatasource.DBInstanceType.DATABASE_SADBEL)
        source.ExecuteNonQuery("INSERT INTO [PLDA IMPORT HEADER] ([Code], [Header], [A1], [A2], [A3], [A4], [A5]) VALUES  ('000000237661242485047', 1, 'IM', 'Y', 'P945304849540810005248', '20081207', 'VP442023')", CDatasource.DBInstanceType.DATABASE_SADBEL)
        source.ExecuteNonQuery("INSERT INTO [PLDA IMPORT HEADER] ([Code], [Header], [A1], [A2], [A3], [A4], [A5]) VALUES  ('000000958309650421142', 1, 'IM', 'Z', 'P945304849540810005250', '20081208', 'VP442026')", CDatasource.DBInstanceType.DATABASE_SADBEL)
        source.ExecuteNonQuery("INSERT INTO [PLDA IMPORT HEADER] ([Code], [Header], [A1], [A2], [A3], [A4], [A5]) VALUES  ('000000372680902481079', 1, 'IM', 'Y', 'P945304849540810005255', '20081209', 'VP442029')", CDatasource.DBInstanceType.DATABASE_SADBEL)
        source.ExecuteNonQuery("INSERT INTO [PLDA IMPORT HEADER] ([Code], [Header], [A1], [A2], [A3], [A4], [A5]) VALUES  ('000000757792890071868', 1, 'IM', 'Z', 'P945304849540810005286', '20081210', 'VP442037')", CDatasource.DBInstanceType.DATABASE_SADBEL)
    End Sub

    <TestCleanup()> Public Sub Cleanup()
        source.ExecuteNonQuery("DELETE FROM [PLDA IMPORT HEADER]", CDatasource.DBInstanceType.DATABASE_SADBEL)
        source = Nothing
    End Sub


    <TestMethod()> Public Sub TestUpdateSingleRecordQueried()
        Dim rstTemp As ADODB.Recordset = New ADODB.Recordset
        Dim rstTest As ADODB.Recordset = New ADODB.Recordset
        Dim success As Integer

        rstTemp = source.ExecuteQuery("SELECT * FROM [PLDA IMPORT HEADER] WHERE [CODE] = '000000237661242485047'", DBInstanceType.DATABASE_SADBEL)
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
            Dim wrapperClass As New CRecordset(rstTemp, rstTemp.Bookmark)
            success = source.UpdateSadbel(wrapperClass, SadbelTableType.PLDA_IMPORT_HEADER)
            Assert.IsTrue(success = 0)

            'Confirm that the changes has been saved to DB
            Threading.Thread.Sleep(1000)
            rstTest = source.ExecuteQuery("SELECT * FROM [PLDA IMPORT HEADER] WHERE [CODE] = '000000237661242485047'", DBInstanceType.DATABASE_SADBEL)
            
            Assert.AreEqual(1, rstTest.RecordCount)
            Assert.AreEqual("20081206", rstTest.Fields("A4").Value)
            Assert.AreEqual("VP442020", rstTest.Fields("A5").Value)

            rstTest.MoveFirst()
            rstTest.Fields("A4").Value = strOriginalA4
            rstTest.Fields("A5").Value = strOriginalA5

            'Revert to original data
            wrapperClass = New CRecordset(rstTest, rstTest.Bookmark)
            success = source.UpdateSadbel(wrapperClass, SadbelTableType.PLDA_IMPORT_HEADER)
            Assert.IsTrue(success = 0)
        End If
    End Sub

    <TestMethod()> Public Sub TestUpdateMultipleRecordQueried()
        Dim rstTemp As ADODB.Recordset = New ADODB.Recordset
        Dim rstTest As ADODB.Recordset = New ADODB.Recordset
        Dim success As Integer

        rstTemp = source.ExecuteQuery("SELECT * FROM [PLDA IMPORT HEADER]", DBInstanceType.DATABASE_SADBEL)
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
            Dim wrapperClass As New CRecordset(rstTemp, rstTemp.Bookmark)
            success = source.UpdateSadbel(wrapperClass, SadbelTableType.PLDA_IMPORT_HEADER)
            Assert.IsTrue(success = 0)

            'Confirm that the changes has been saved to DB
            Threading.Thread.Sleep(1000)
            rstTest = source.ExecuteQuery("SELECT * FROM [PLDA IMPORT HEADER] WHERE [CODE] = '" & strCode & "'", DBInstanceType.DATABASE_SADBEL)
            
            Assert.AreEqual(1, rstTest.RecordCount)
            Assert.AreEqual("20081299", rstTest.Fields("A4").Value)
            Assert.AreEqual("VP442027", rstTest.Fields("A5").Value)

            rstTest.MoveFirst()
            rstTest.Fields("A4").Value = strOriginalA4
            rstTest.Fields("A5").Value = strOriginalA5

            'Revert to original data 
            wrapperClass = New CRecordset(rstTest, rstTest.Bookmark)
            success = source.UpdateSadbel(wrapperClass, SadbelTableType.PLDA_IMPORT_HEADER)
            Assert.IsTrue(success = 0)
        End If
    End Sub

    <TestMethod()> Public Sub TestUpdateMultipleRecordQueriedWithADOFilter()
        Dim rstTemp As ADODB.Recordset = New ADODB.Recordset
        Dim rstTest As ADODB.Recordset = New ADODB.Recordset
        Dim success As Integer

        rstTemp = source.ExecuteQuery("SELECT * FROM [PLDA IMPORT HEADER]", DBInstanceType.DATABASE_SADBEL)
        
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
            Dim wrapperClass As New CRecordset(rstTemp, rstTemp.Bookmark)
            success = source.UpdateSadbel(wrapperClass, SadbelTableType.PLDA_IMPORT_HEADER)
            Assert.IsTrue(success = 0)

            'Confirm that the changes has been saved to DB
            Threading.Thread.Sleep(1000)
            rstTest = source.ExecuteQuery("SELECT * FROM [PLDA IMPORT HEADER] WHERE [CODE] = '" & strCode & "'", DBInstanceType.DATABASE_SADBEL)

            Assert.AreEqual(1, rstTest.RecordCount)
            Assert.AreEqual("20081277", rstTest.Fields("A4").Value)
            Assert.AreEqual("VP442055", rstTest.Fields("A5").Value)

            rstTest.MoveFirst()
            rstTest.Fields("A4").Value = strOriginalA4
            rstTest.Fields("A5").Value = strOriginalA5

            'Revert to original data
            wrapperClass = New CRecordset(rstTest, rstTest.Bookmark)
            success = source.UpdateSadbel(wrapperClass, SadbelTableType.PLDA_IMPORT_HEADER)
            Assert.IsTrue(success = 0)
        End If
    End Sub

    '<TestMethod()> Public Sub TestInsert()
    '    Dim source As CDatasource = New CDatasource

    '    Dim conADO As ADODB.Connection = New ADODB.Connection
    '    Dim rstTemp As ADODB.Recordset = New ADODB.Recordset
    '    Dim rstTest As ADODB.Recordset = New ADODB.Recordset
    '    Dim success As Integer

    '    ConnectDB(conADO, My.Application.Info.DirectoryPath, "mdb_sadbel.mdb")
    '    Dim strCon As String = conADO.ConnectionString
    '    RstOpen("SELECT * FROM [PLDA IMPORT HEADER]", conADO, rstTemp, adOpenKeyset, adLockOptimistic, , True)
    '    DisconnectDB(conADO)

    '    Assert.AreEqual(5, rstTemp.RecordCount)

    '    If rstTemp.RecordCount > 0 Then
    '        'Add new record to Recordset
    '        rstTemp.AddNew()
    '        'Assert.AreEqual(6, rstTemp.RecordCount)

    '        'Insert Values
    '        rstTemp.Fields("Code").Value() = "000000993949532508811"
    '        rstTemp.Fields("Header").Value() = "1"
    '        rstTemp.Fields("A1").Value() = "IM"
    '        rstTemp.Fields("A2").Value() = "Z"
    '        rstTemp.Fields("A3").Value() = "P945304849540810005111"
    '        rstTemp.Fields("A4").Value() = "20081288"
    '        rstTemp.Fields("A5").Value() = "VP442069"

    '        'Use CubelibDatasource Update method
    '        Dim wrapperClass As New CRecordset(rstTemp, rstTemp.Bookmark)
    '        success = source.UpdateSadbel(wrapperClass, SadbelTableType.PLDA_IMPORT_HEADER)
    '        Assert.IsTrue(success = 0)

    '        'Confirm that the changes has been saved to DB
    '        ConnectDB(conADO, My.Application.Info.DirectoryPath, "mdb_sadbel.mdb")
    '        RstOpen("SELECT * FROM [PLDA IMPORT HEADER] WHERE [CODE] = '000000993949532508811'", conADO, rstTest, adOpenKeyset, adLockOptimistic, , True)
    '        DisconnectDB(conADO)

    '        Assert.AreEqual(1, rstTest.RecordCount)
    '        Assert.AreEqual(Convert.ToDouble(1), rstTest.Fields("Header").Value)
    '        Assert.AreEqual("IM", rstTest.Fields("A1").Value)
    '        Assert.AreEqual("Z", rstTest.Fields("A2").Value)
    '        Assert.AreEqual("P945304849540810005111", rstTest.Fields("A3").Value)
    '        Assert.AreEqual("20081288", rstTest.Fields("A4").Value)
    '        Assert.AreEqual("VP442069", rstTest.Fields("A5").Value)

    '        'Remove inserted data 
    '        success = source.ExecuteNonQuery("DELETE FROM [PLDA IMPORT HEADER] WHERE [CODE] = '000000993949532508811'", CDatasource.DBInstanceType.DATABASE_SADBEL)
    '        Assert.IsTrue(success = 0)

    '        RstClose(rstTemp)
    '        RstClose(rstTest)
    '    End If
    'End Sub
End Class

