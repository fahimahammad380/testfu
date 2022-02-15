Imports MySql.Data.MySqlClient
Public Class DTConnection
    Dim conect As New MySqlConnection("server='107.180.26.71';port='3306';user='appdev';password='F1l@T3ch#8062$';database='apptimecard'")
    'Dim conect As New MySqlConnection("server='localhost';user='root';password='';database='filasco_db'")
    Public Function open() As MySqlConnection
        Try
            conect.Open()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        Return conect
    End Function

    Public Function close() As MySqlConnection
        Try
            conect.Close()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        Return conect
    End Function
End Class