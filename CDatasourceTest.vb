﻿Imports System.Text
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports CubeLibDataSource
Imports ADODB.CursorTypeEnum
Imports ADODB.LockTypeEnum
Imports CubeLibDataSource.CDatasource
Imports System.IO
Imports System.IO.Compression
Imports System.Drawing

<TestClass()> Public Class CDatasourceTest

    Private source As CDatasource

    <TestInitialize()> Public Sub Init()
        source = New CDatasource
        source.Open(AppDomain.CurrentDomain.BaseDirectory)

        source.ExecuteNonQuery("INSERT INTO [PLDA IMPORT HEADER] ([Code], [Header], [A1], [A2], [A3], [A4], [A5]) VALUES  ('000000993949532508849', 1, 'IM', 'Z', 'P945304849540810005246', '20081206', 'VP442020')", CDatasource.DBInstanceType.DATABASE_SADBEL)
        source.ExecuteNonQuery("INSERT INTO [PLDA IMPORT HEADER] ([Code], [Header], [A1], [A2], [A3], [A4], [A5]) VALUES  ('000000237661242485047', 1, 'IM', 'Y', 'P945304849540810005248', '20081207', 'VP442023')", CDatasource.DBInstanceType.DATABASE_SADBEL)
        source.ExecuteNonQuery("INSERT INTO [PLDA IMPORT HEADER] ([Code], [Header], [A1], [A2], [A3], [A4], [A5]) VALUES  ('000000958309650421142', 1, 'IM', 'Z', 'P945304849540810005250', '20081208', 'VP442026')", CDatasource.DBInstanceType.DATABASE_SADBEL)
        source.ExecuteNonQuery("INSERT INTO [PLDA IMPORT HEADER] ([Code], [Header], [A1], [A2], [A3], [A4], [A5]) VALUES  ('000000372680902481079', 1, 'IM', 'Y', 'P945304849540810005255', '20081209', 'VP442029')", CDatasource.DBInstanceType.DATABASE_SADBEL)
        source.ExecuteNonQuery("INSERT INTO [PLDA IMPORT HEADER] ([Code], [Header], [A1], [A2], [A3], [A4], [A5]) VALUES  ('000000757792890071868', 1, 'IM', 'Z', 'P945304849540810005286', '20081210', 'VP442037')", CDatasource.DBInstanceType.DATABASE_SADBEL)

        source.ExecuteNonQuery("INSERT INTO [DATA_NCTS] ([DATA_NCTS_MSG_ID], [CODE], [LOGID DESCRIPTION], [TYPE], [COMM], [USER NO], [LAST MODIFIED BY]) VALUES  (678, '054524877079401395030', 'DHL', 'T', 'S', 3, 'Olivier')", CDatasource.DBInstanceType.DATABASE_EDIFACT)

        source.ExecuteNonQuery("INSERT INTO [Tree] ([LEVEL], [PARENT ID], [TREE ID], [ROOT ID], [DESCRIPTION], [IMAGE], [PICTURE]) VALUES  (1, 'ROOT', '2014', '2014', 'Root for 2014', 1, 1)", CDatasource.DBInstanceType.DATABASE_REPERTORY)
    End Sub

    <TestCleanup()> Public Sub Cleanup()
        source.ExecuteNonQuery("DELETE FROM [PLDA IMPORT HEADER]", CDatasource.DBInstanceType.DATABASE_SADBEL)
        source.ExecuteNonQuery("DELETE FROM [DATA_NCTS]", CDatasource.DBInstanceType.DATABASE_EDIFACT)
        source.ExecuteNonQuery("DELETE FROM [Tree]", CDatasource.DBInstanceType.DATABASE_REPERTORY)
        source.Dispose()
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
            success = source.UpdateSadbel(rstTemp, SadbelTableType.PLDA_IMPORT_HEADER)
            Assert.IsTrue(success = 1)

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
            success = source.UpdateSadbel(rstTest, SadbelTableType.PLDA_IMPORT_HEADER)
            Assert.IsTrue(success = 1)
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
            success = source.UpdateSadbel(rstTemp, SadbelTableType.PLDA_IMPORT_HEADER)
            Assert.IsTrue(success = 1)

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
            success = source.UpdateSadbel(rstTest, SadbelTableType.PLDA_IMPORT_HEADER)
            Assert.IsTrue(success = 1)
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
            success = source.UpdateSadbel(rstTemp, SadbelTableType.PLDA_IMPORT_HEADER)
            Assert.IsTrue(success = 1)

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
            success = source.UpdateSadbel(rstTest, SadbelTableType.PLDA_IMPORT_HEADER)
            Assert.IsTrue(success = 1)
        End If
    End Sub

    <TestMethod()> Public Sub TestInsert()
        Dim rstTemp As ADODB.Recordset = New ADODB.Recordset
        Dim rstTest As ADODB.Recordset = New ADODB.Recordset
        Dim success As Integer

        rstTemp = source.ExecuteQuery("SELECT * FROM [PLDA IMPORT HEADER]", DBInstanceType.DATABASE_SADBEL)

        Assert.AreEqual(5, rstTemp.RecordCount)

        If rstTemp.RecordCount > 0 Then
            'Add new record to Recordset
            rstTemp.AddNew()

            'Insert Values
            Dim strCode As String = "000000993949532508811"
            rstTemp.Fields("Code").Value() = strCode
            rstTemp.Fields("Header").Value() = "1"
            rstTemp.Fields("A1").Value() = "IM"
            rstTemp.Fields("A2").Value() = "Z"
            rstTemp.Fields("A3").Value() = "P945304849540810005111"
            rstTemp.Fields("A4").Value() = "20081288"
            rstTemp.Fields("A5").Value() = "VP442069"

            'Use CubelibDatasource Update method
            success = source.InsertSadbel(rstTemp, SadbelTableType.PLDA_IMPORT_HEADER)
            Assert.IsTrue(success = 0)

            'Confirm that the changes has been saved to DB
            Threading.Thread.Sleep(1000)
            rstTest = source.ExecuteQuery("SELECT * FROM [PLDA IMPORT HEADER] WHERE [CODE] = '" & strCode & "'", DBInstanceType.DATABASE_SADBEL)


            Assert.AreEqual(1, rstTest.RecordCount)
            Assert.AreEqual(Convert.ToDouble(1), rstTest.Fields("Header").Value)
            Assert.AreEqual("IM", rstTest.Fields("A1").Value)
            Assert.AreEqual("Z", rstTest.Fields("A2").Value)
            Assert.AreEqual("P945304849540810005111", rstTest.Fields("A3").Value)
            Assert.AreEqual("20081288", rstTest.Fields("A4").Value)
            Assert.AreEqual("VP442069", rstTest.Fields("A5").Value)

            'Remove inserted data 
            success = source.ExecuteNonQuery("DELETE FROM [PLDA IMPORT HEADER] WHERE [CODE] = '" & strCode & "'", CDatasource.DBInstanceType.DATABASE_SADBEL)
            Assert.IsTrue(success = 1)
        End If
    End Sub

    <TestMethod()> Public Sub TestGetEnumFromTableName()
        Dim result As Integer

        result = source.GetEnumFromTableName(SadbelTableType.AUTHORIZEDPARTIES.ToString, DBInstanceType.DATABASE_SADBEL)
        Assert.IsTrue(result = 0)

        result = source.GetEnumFromTableName(DataTableType.MASTER.ToString, DBInstanceType.DATABASE_DATA)
        Assert.IsTrue(result = 1)

        result = source.GetEnumFromTableName(EdiHistoryTableType.DATA_NCTS_BERICHT_DOUANEKANTOOR.ToString, DBInstanceType.DATABASE_EDI_HISTORY)
        Assert.IsTrue(result = 3)

        result = source.GetEnumFromTableName(EdifactTableType.DATA_NCTS_BERICHT_VERVOER.ToString, DBInstanceType.DATABASE_EDIFACT)
        Assert.IsTrue(result = 6)

        result = source.GetEnumFromTableName(SadbelHistoryTableType.IMPORT_HEADER.ToString, DBInstanceType.DATABASE_HISTORY)
        Assert.IsTrue(result = 16)

        result = source.GetEnumFromTableName(RepertoryTableType.Fields.ToString, DBInstanceType.DATABASE_REPERTORY)
        Assert.IsTrue(result = 4)

        result = source.GetEnumFromTableName(SchedulerTableType.Archiver_Properties.ToString, DBInstanceType.DATABASE_SCHEDULER)
        Assert.IsTrue(result = 0)

        result = source.GetEnumFromTableName(TaricTableType.COMMON.ToString, DBInstanceType.DATABASE_TARIC)
        Assert.IsTrue(result = 2)

        result = source.GetEnumFromTableName(TemplateCPTableType.DBProps.ToString, DBInstanceType.DATABASE_TEMPLATE)
        Assert.IsTrue(result = 10)
    End Sub

    <TestMethod()> Public Sub TestSelectEdifact()
        Dim rstTemp As ADODB.Recordset

        rstTemp = source.ExecuteQuery("SELECT TOP 1 * FROM [DATA_NCTS]", DBInstanceType.DATABASE_EDIFACT)
        Assert.AreEqual(1, rstTemp.RecordCount)
    End Sub

    <TestMethod()> Public Sub TestUpdateSingleRecordQueriedEdifactNoPK()
        Dim rstTemp As ADODB.Recordset = New ADODB.Recordset
        Dim rstTest As ADODB.Recordset = New ADODB.Recordset
        Dim success As Integer

        rstTemp = source.ExecuteQuery("SELECT * FROM [BOX_SEARCH_MAP] WHERE [BOX CODE] = 'AG' AND [NCTS_IEM_TMS_ID] = 316", DBInstanceType.DATABASE_EDIFACT)
        Assert.AreEqual(1, rstTemp.RecordCount)

        If rstTemp.RecordCount > 0 Then
            rstTemp.MoveFirst()

            'Store Original Values
            Dim iemID As Integer = rstTemp.Fields("NCTS_IEM_ID").Value
            Dim boxCode As String = rstTemp.Fields("BOX CODE").Value
            Dim iemTmsID As Integer = rstTemp.Fields("NCTS_IEM_TMS_ID").Value

            'Edit A4 and A5
            rstTemp.Fields("NCTS_IEM_ID").Value = 4

            'Use CubelibDatasource Update method
            success = source.UpdateEdifact(rstTemp, EdifactTableType.BOX_SEARCH_MAP)

            'Update did not succeed because table has no primary key
            Assert.IsTrue(success = -1)
        End If
    End Sub

    <TestMethod()> Public Sub TestUpdateSingleRecordQueriedEdifact()
        Dim rstTemp As ADODB.Recordset = New ADODB.Recordset
        Dim rstTest As ADODB.Recordset = New ADODB.Recordset
        Dim success As Integer

        rstTemp = source.ExecuteQuery("SELECT * FROM [DATA_NCTS] WHERE [CODE] = '054524877079401395030'", DBInstanceType.DATABASE_EDIFACT)
        Assert.AreEqual(1, rstTemp.RecordCount)

        If rstTemp.RecordCount > 0 Then
            rstTemp.MoveFirst()

            'Store Original Values
            Dim code As String = rstTemp.Fields("CODE").Value
            Dim logID As String = rstTemp.Fields("LOGID DESCRIPTION").Value
            Dim type As String = rstTemp.Fields("TYPE").Value

            'Edit LOGID DESCRIPTION and TYPE
            rstTemp.Fields("LOGID DESCRIPTION").Value = "LBC"
            rstTemp.Fields("TYPE").Value = "T"

            'Use CubelibDatasource Update method
            success = source.UpdateEdifact(rstTemp, EdifactTableType.DATA_NCTS)
            Assert.IsTrue(success = 1)

            'Confirm that the changes has been saved to DB
            Threading.Thread.Sleep(1000)
            rstTest = source.ExecuteQuery("SELECT * FROM [DATA_NCTS] WHERE [CODE] = '" & code & "'", DBInstanceType.DATABASE_EDIFACT)

            Assert.AreEqual(1, rstTest.RecordCount)
            Assert.AreEqual("LBC", rstTest.Fields("LOGID DESCRIPTION").Value)
            Assert.AreEqual("T", rstTest.Fields("TYPE").Value)

            rstTest.MoveFirst()
            rstTest.Fields("LOGID DESCRIPTION").Value = logID
            rstTest.Fields("TYPE").Value = type

            'Revert to original data
            success = source.UpdateEdifact(rstTest, EdifactTableType.DATA_NCTS)
            Assert.IsTrue(success = 1)
        End If
    End Sub

    <TestMethod()> Public Sub TestSelectRepertory()
        Dim rstTemp As ADODB.Recordset

        rstTemp = source.ExecuteQuery("SELECT TOP 1 * FROM [Tree]", DBInstanceType.DATABASE_REPERTORY)
        Assert.AreEqual(1, rstTemp.RecordCount)
    End Sub

    <TestMethod()> Public Sub TestInsertIdentity()
        Dim rstTemp As ADODB.Recordset = New ADODB.Recordset
        Dim rstTest As ADODB.Recordset = New ADODB.Recordset
        Dim identity As Integer

        rstTemp = source.ExecuteQuery("SELECT * FROM [AuthorizedParties] WHERE 1=0", DBInstanceType.DATABASE_SADBEL)

        'Add new record to Recordset
        rstTemp.AddNew()

        'Insert Values
        rstTemp.Fields("Auth_Name").Value() = "CANDS"

        'Use CubelibDatasource Update method
        identity = source.InsertSadbel(rstTemp, SadbelTableType.AUTHORIZEDPARTIES)
        Assert.IsTrue(identity = 1)

        rstTemp = source.ExecuteQuery("SELECT * FROM [AuthorizedParties] WHERE 1=0", DBInstanceType.DATABASE_SADBEL)

        'Add new record to Recordset
        rstTemp.AddNew()

        'Insert Values
        rstTemp.Fields("Auth_Name").Value() = "CANDS"

        'Use CubelibDatasource Update method
        identity = source.InsertSadbel(rstTemp, SadbelTableType.AUTHORIZEDPARTIES)
        Assert.IsTrue(identity = 2)

    End Sub

    <TestMethod()> Public Sub TestCreateDropTable()
        Dim success As Integer

        success = source.ExecuteNonQuery("CREATE TABLE TEST_TABLE (FirstName CHAR, LastName CHAR)", DBInstanceType.DATABASE_SADBEL)
        Assert.AreEqual(0, success)

        success = source.ExecuteNonQuery("DROP TABLE TEST_TABLE", DBInstanceType.DATABASE_SADBEL)
        Assert.AreEqual(0, success)
    End Sub

    <TestMethod()> Public Sub TestInsertPLDAMessages()
        Dim rstTemp As ADODB.Recordset = New ADODB.Recordset
        Dim rstTest As ADODB.Recordset = New ADODB.Recordset
        Dim identity As Integer

        rstTemp = source.ExecuteQuery("SELECT * FROM [PLDA MESSAGES] WHERE 1=0", DBInstanceType.DATABASE_SADBEL)

        'Add new record to Recordset
        rstTemp.AddNew()

        'Insert Values
        rstTemp.Fields("Code").Value() = "000000092479884624482"
        rstTemp.Fields("DType").Value = 18
        rstTemp.Fields("Message_Date").Value = Now
        rstTemp.Fields("Message_StatusType").Value = "Document"
        rstTemp.Fields("User_ID").Value = 1
        rstTemp.Fields("Message_Reference").Value = "1402000377961000106278"

        'Use CubelibDatasource Update method
        identity = source.InsertSadbel(rstTemp, SadbelTableType.PLDA_MESSAGES)
        Assert.IsTrue(identity = 1)

    End Sub

    <TestMethod()> Public Sub TestDefaultViewColumns()
        Dim rstTemp As ADODB.Recordset

        rstTemp = source.ExecuteQuery("SELECT TOP 1 * FROM [DefaultViewColumns]", DBInstanceType.DATABASE_TEMPLATE)
        Assert.AreEqual(1, rstTemp.RecordCount)
    End Sub

    '<TestMethod()> Public Sub TestUVCFormatCondition()
    '    Dim rstTemp As ADODB.Recordset
    '    Dim objTest As New FTest

    '    rstTemp = source.ExecuteQuery("SELECT TOP 1 * FROM [UVCFormatCondition] WHERE FC_ID = 2032", DBInstanceType.DATABASE_TEMPLATE)
    '    Assert.AreEqual(1, rstTemp.RecordCount)

    '    MsgBox(Image.FromFile("save.jpeg"))
    '    objTest.PictureBox1.Image = Image.FromFile("save.jpeg")
    '    objTest.Show()
    '    'objTest = Nothing
    'End Sub
End Class