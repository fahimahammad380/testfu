
Public Class Form2
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim inputTagCollection = wb.Document.GetElementsByTagName("input")
        For Each inputTag As HtmlElement In inputTagCollection
            If inputTag.OuterHtml.Contains("post") Then
                inputTag.InvokeMember("click")
            End If
        Next
    End Sub

    Private Sub Form2_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        wb.Navigate("voip.forumfree.it/?act=idx")
    End Sub
End Class