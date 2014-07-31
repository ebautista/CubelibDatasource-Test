Imports System.Text
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports CubelibDatasource
Imports CubelibDatasource.CDatasource

<TestClass()> Public Class ExceptionTests

    Private source As CDatasource = New CDatasource

    <TestMethod()> Public Sub TestErrorPersistencePath()
        Try
            source.SetPersistencePath(vbNullString)
            Assert.Fail()
        Catch ex As Exception
            Debug.Print(ex.Message)
        End Try

    End Sub

    <TestMethod()> Public Sub TestErrorDatabasePropertyNotInitializedNonQuery()
        Try
            source.ExecuteNonQuery("DROP TABLE TEST_TABLE", DBInstanceType.DATABASE_SADBEL)
            Assert.Fail()
        Catch ex As Exception
            Debug.Print(ex.Message)
        End Try

    End Sub

    <TestMethod()> Public Sub TestErrorDatabasePropertyNotInitializedQuery()
        Try
            source.ExecuteQuery("SELECT * FROM TEST", DBInstanceType.DATABASE_SADBEL)
            Assert.Fail()
        Catch ex As Exception
            Debug.Print(ex.Message)
        End Try

    End Sub

End Class