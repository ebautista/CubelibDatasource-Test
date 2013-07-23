Imports System.Text
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports CubelibDatasource
Imports ADODB.CursorTypeEnum
Imports ADODB.LockTypeEnum

<TestClass()> Public Class CDatasourceTest

    <TestMethod()> Public Sub TestUpdate()
        Dim source As CDatasource = New CDatasource

        Dim conADO As ADODB.Connection = New ADODB.Connection
        Dim rstTemp As ADODB.Recordset = New ADODB.Recordset
        Dim rstTest As ADODB.Recordset = New ADODB.Recordset
        Dim success As Integer
        
        ConnectDB(conADO, My.Application.Info.DirectoryPath, "mdb_sadbel.mdb")

        RstOpen("SELECT * FROM [PLDA IMPORT HEADER] WHERE [CODE] = '000000993949532508849'", conADO, rstTemp, adOpenKeyset, adLockOptimistic)

        If rstTemp.RecordCount > 0 Then
            rstTemp.MoveFirst()
            rstTemp.Fields("A4").Value = "20081206"
            rstTemp.Fields("A5").Value = "VP442020"

            success = source.Update(rstTemp, rstTemp.Bookmark)
            Assert.IsTrue(success = 0)

            RstOpen("SELECT * FROM [PLDA IMPORT HEADER] WHERE [CODE] = '000000993949532508849'", conADO, rstTest, adOpenKeyset, adLockOptimistic)
            Assert.AreEqual("20081206", rstTest.Fields("A4").Value)
            Assert.AreEqual("VP442020", rstTest.Fields("A5").Value)

            RstClose(rstTemp)
            RstClose(rstTest)
        End If
    End Sub

End Class

