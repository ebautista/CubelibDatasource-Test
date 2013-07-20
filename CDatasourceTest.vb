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
        Dim fldTemp As ADODB.Field

        ConnectDB(conADO, "C:\ClearingPoint\Data", "mdb_sadbel.mdb")

        RstOpen("SELECT * FROM [PLDA IMPORT HEADER]", conADO, rstTemp, adOpenKeyset, adLockOptimistic)

        For Each fldTemp In rstTemp.Fields
            Debug.Print(fldTemp.Name & " - " & fldTemp.Value)
        Next fldTemp
    End Sub

End Class

