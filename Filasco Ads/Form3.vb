Option Explicit On
Imports System.Net
Imports MySql.Data.MySqlClient
Public Class Form3
#Region "declare"
    Dim mycmd As New MySqlCommand
    Dim myconnection As New DTConnection
    Dim objreader As MySqlDataReader
    Dim Mysqlconn As MySqlConnection
#End Region
    'for sleep
    'ivate Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
    'for sleep
    Dim InitialZoom As Integer = 100
    Public Enum Exec
        OLECMDID_OPTICAL_ZOOM = 63
    End Enum
    Private Enum execOpt
        OLECMDEXECOPT_DODEFAULT = 0
        OLECMDEXECOPT_PROMPTUSER = 1
        OLECMDEXECOPT_DONTPROMPTUSER = 2
        OLECMDEXECOPT_SHOWHELP = 3
    End Enum
    Public Sub PerformZoom(ByVal Value As Integer)
        Try
            Dim Res As Object = Nothing
            Dim MyWeb As Object
            MyWeb = Me.callsMarket.ActiveXInstance
            MyWeb.ExecWB(Exec.OLECMDID_OPTICAL_ZOOM, execOpt.OLECMDEXECOPT_PROMPTUSER, CObj(Value), CObj(IntPtr.Zero))
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

        Try
            Dim Res As Object = Nothing
            Dim MyWeb As Object
            MyWeb = Me.calltermination.ActiveXInstance
            MyWeb.ExecWB(Exec.OLECMDID_OPTICAL_ZOOM, execOpt.OLECMDEXECOPT_PROMPTUSER, CObj(Value), CObj(IntPtr.Zero))
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

        Try
            Dim Res As Object = Nothing
            Dim MyWeb As Object
            MyWeb = Me.voiphelp.ActiveXInstance
            MyWeb.ExecWB(Exec.OLECMDID_OPTICAL_ZOOM, execOpt.OLECMDEXECOPT_PROMPTUSER, CObj(Value), CObj(IntPtr.Zero))
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

        Try
            Dim Res As Object = Nothing
            Dim MyWeb As Object
            MyWeb = Me.bestVoip.ActiveXInstance
            MyWeb.ExecWB(Exec.OLECMDID_OPTICAL_ZOOM, execOpt.OLECMDEXECOPT_PROMPTUSER, CObj(Value), CObj(IntPtr.Zero))
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        Try
            Dim Res As Object = Nothing
            Dim MyWeb As Object
            MyWeb = Me.forumVoip.ActiveXInstance
            MyWeb.ExecWB(Exec.OLECMDID_OPTICAL_ZOOM, execOpt.OLECMDEXECOPT_PROMPTUSER, CObj(Value), CObj(IntPtr.Zero))
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        Try
            Dim Res As Object = Nothing
            Dim MyWeb As Object
            MyWeb = Me.voiptraffic.ActiveXInstance
            MyWeb.ExecWB(Exec.OLECMDID_OPTICAL_ZOOM, execOpt.OLECMDEXECOPT_PROMPTUSER, CObj(Value), CObj(IntPtr.Zero))
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        InitialZoom += 10
        PerformZoom(InitialZoom)
    End Sub
    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        InitialZoom -= 10
        PerformZoom(InitialZoom)
    End Sub
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles btnFit.Click
        InitialZoom = 40
        PerformZoom(InitialZoom)
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles btnSetData.Click
        Try
            callsMarket.Document.GetElementById("subject").SetAttribute("value", txtSubject.Text)
            callsMarket.Document.GetElementById("message").SetAttribute("value", txtBody.Text)

            calltermination.Document.GetElementById("subject").SetAttribute("value", txtSubject.Text)
            calltermination.Document.GetElementById("message").SetAttribute("value", txtBody.Text)

            forumVoip.Document.GetElementById("subject").SetAttribute("value", txtSubject.Text)
            forumVoip.Document.GetElementById("message").SetAttribute("value", txtBody.Text)

            voiphelp.Document.GetElementById("subject").SetAttribute("value", txtSubject.Text)
            voiphelp.Document.GetElementById("message").SetAttribute("value", txtBody.Text)

            Dim inputTagCollection3 = bestVoip.Document.GetElementsByTagName("input")
            For Each inputTag As HtmlElement In inputTagCollection3
                If inputTag.OuterHtml.Contains("subject") Then
                    inputTag.SetAttribute("value", txtSubject.Text)
                End If
            Next
            bestVoip.Document.GetElementById("message").SetAttribute("value", txtBody.Text)

            voiptraffic.Document.GetElementById("subject").SetAttribute("value", txtSubject.Text)

            Dim inputTagCollection_voiptraffic = voiptraffic.Document.GetElementsByTagName("input")
            For Each inputTag As HtmlElement In inputTagCollection_voiptraffic
                If inputTag.OuterHtml.Contains("subject") Then
                    inputTag.SetAttribute("value", txtSubject.Text)
                End If
            Next
            Dim x = voiptraffic.Document.GetElementById("subject")
        Catch ex As Exception
            'MsgBox(ex.ToString)
        End Try
    End Sub
    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles goBack.Click
        voiptraffic.GoBack()
    End Sub
    Private Sub Button1_Click_2(sender As Object, e As EventArgs) Handles goForward.Click
        voiptraffic.GoForward()
    End Sub
    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles btnReload.Click
        On Error Resume Next
        reload()
        callsMarket.Navigate("https://www.callsmarket.com/posting.php?mode=post&f=1")
        calltermination.Navigate("http://www.calltermination.com/posting.php?f=6&mode=post&sid=2a345099e9b17451653d221be18850e1")
        forumVoip.Navigate("http://www.forumvoip.com/posting.php?mode=newtopic&f=3")
        bestVoip.Navigate("https://bestvoip.forumotion.com/post?f=3&mode=newtopic")
        voiptraffic.Navigate("https://voiptraffic.forumotion.com/post?f=7&mode=newtopic")
        voiphelp.Navigate("http://voiphelp24.com/posting.php?mode=post&f=4")
        goBack.PerformClick()
    End Sub
    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        Timer1.Start()
    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        Label1.Text = Val(Label1.Text) + 1
        If Label1.Text = 5 Then
            ComboBoxSubject.SelectedIndex = 0
        ElseIf Label1.Text = 100 Then
            ComboBoxSubject.SelectedIndex = 1
        ElseIf Label1.Text = 200 Then
            ComboBoxSubject.SelectedIndex = 2
        ElseIf Label1.Text = 300 Then
            ComboBoxSubject.SelectedIndex = 3
        ElseIf Label1.Text = 400 Then
            Label1.Text = 1
        End If
        ProgressBar1.Value += 1
        If ProgressBar1.Value = 5 Then
            Timer1.Interval = 100 '0
            btnFit.PerformClick()
            btnReload.PerformClick()
        ElseIf ProgressBar1.Value = 18 Then
            btnSetData.PerformClick()
        ElseIf ProgressBar1.Value = 20 Then
            'Timer1.Dispose()
            btnPost.PerformClick()
            Timer1.Interval = 100 '00
        ElseIf ProgressBar1.Value = 100 Then
            ProgressBar1.Value = 1
        End If
    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBoxSubject.SelectedIndexChanged
        If ComboBoxSubject.Text = "Algeria,Albania,Afghan,Armenia,BD,Belarus,Chad,Cuba,Camerron" Then
            txtSubject.Text = "Algeria,Albania,Afghan,Armenia,BD,Belarus,Chad,Cuba,Camerron"
            txtBody.Text = "Dear Provider,

Greetings from Filasco America LLC!

we are offering Long term business opportunity for direct non-CLI route providers(A-Z destinations).
We are Looking for Stable/High Quality routes for Pure Retail Traffic.

Wholesale Route Providers also welcome.

Urgently Need below destinations for Pure Retail Traffic

Algeria
Albania
Afghan
Armenia
Bangladesh
Belarus
Chad
Cuba
Camerron

Thanks.

Best regards
Jacob Hoffman
Sales Executive Officer at Filasco America LLC.
Skype: live:jacob.hoffman_8
Email :jacob.hoffman@filasco.net
Website: http://www.filasco.net
Phone: +1(646) 691-2952
Fax: +1(866) 996-8535
Corporate Office: 47th street, Suite B3, Brooklyn, NY 11220"
        ElseIf ComboBoxSubject.Text = "DRC,Ethiopia,Ghana,Gambia,Haiti,Indonesia,Iran,Iraq,Jamaica" Then
            txtSubject.Text = "DRC,Ethiopia,Ghana,Gambia,Haiti,Indonesia,Iran,Iraq,Jamaica"
            txtBody.Text = "Dear Provider,

Greetings from Filasco America LLC!

we are offering Long term business opportunity for direct non-CLI route providers(A-Z destinations).
We are Looking for Stable/High Quality routes for Pure Retail Traffic.

Wholesale Route Providers also welcome.

Urgently Need below destinations for Pure Retail Traffic

DRC
Ethiopia
Ghana
Gambia
Haiti
Indonesia
Iran
Iraq
Jamaica

Thanks.

Best regards
Jacob Hoffman
Sales Executive Officer at Filasco America LLC.
Skype: live:jacob.hoffman_8
Email :jacob.hoffman@filasco.net
Website: http://www.filasco.net
Phone: +1(646) 691-2952
Fax: +1(866) 996-8535
Corporate Office: 47th street, Suite B3, Brooklyn, NY 11220"
        ElseIf ComboBoxSubject.Text = "Libya,Mali,Mozambique,Mauritania,Morocco,Nigeria,Nepal,Pak" Then
            txtSubject.Text = "Libya,Mali,Mozambique,Mauritania,Morocco,Nigeria,Nepal,Pak"
            txtBody.Text = "Dear Provider,

Greetings from Filasco America LLC!

we are offering Long term business opportunity for direct non-CLI route providers(A-Z destinations).
We are Looking for Stable/High Quality routes for Pure Retail Traffic.

Wholesale Route Providers also welcome.

Urgently Need below destinations for Pure Retail Traffic

Libya
Mali
Mozambique
Mauritania
Morocco
Nigeria
Nepal
Pak

Thanks.

Best regards
Jacob Hoffman
Sales Executive Officer at Filasco America LLC.
Skype: live:jacob.hoffman_8
Email :jacob.hoffman@filasco.net
Website: http://www.filasco.net
Phone: +1(646) 691-2952
Fax: +1(866) 996-8535
Corporate Office: 47th street, Suite B3, Brooklyn, NY 11220"
        ElseIf ComboBoxSubject.Text = "Russia,Spain,SA,Sudan,Turkey,Tunisia,Uganda,UAE,Yemen,Zambia" Then
            txtSubject.Text = "Russia,Spain,SA,Sudan,Turkey,Tunisia,Uganda,UAE,Yemen,Zambia"
            txtBody.Text = "Dear Provider,

Greetings from Filasco America LLC!

we are offering Long term business opportunity for direct non-CLI route providers(A-Z destinations).
We are Looking for Stable/High Quality routes for Pure Retail Traffic.

Wholesale Route Providers also welcome.

Urgently Need below destinations for Pure Retail Traffic

Russia
Spain
South Africa
Sudan
Turkey
Tunisia
Uganda
UAE
Yemen
Zambia

Thanks.

Best regards
Jacob Hoffman
Sales Executive Officer at Filasco America LLC.
Skype: live:jacob.hoffman_8
Email :jacob.hoffman@filasco.net
Website: http://www.filasco.net
Phone: +1(646) 691-2952
Fax: +1(866) 996-8535
Corporate Office: 47th street, Suite B3, Brooklyn, NY 11220"
        End If

        mycmd.Connection = myconnection.open
        mycmd.CommandText = "SELECT * from attempt_to_registration where User_Name = '" + ComboBoxSubject.Text + "' "
        objreader = mycmd.ExecuteReader
        While objreader.Read
            Dim comments As String = objreader.GetString("Comments")
            If ComboBoxSubject.SelectedIndex = 0 Then
                txtBody.Text = comments
                txtSubject.Text = ComboBoxSubject.Text
            ElseIf ComboBoxSubject.SelectedIndex = 1 Then
                txtBody.Text = comments
                txtSubject.Text = ComboBoxSubject.Text
            ElseIf ComboBoxSubject.SelectedIndex = 2 Then
                txtBody.Text = comments
                txtSubject.Text = ComboBoxSubject.Text
            ElseIf ComboBoxSubject.SelectedIndex = 3 Then
                txtBody.Text = comments
                txtSubject.Text = ComboBoxSubject.Text
            End If
        End While
        myconnection.close()
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.HorizontalScroll.Maximum = 0
        Me.AutoScroll = True
        ' On Error Resume Next
        forumVoip.Navigate("http://www.forumvoip.com/posting.php?mode=newtopic&f=3")
        bestVoip.Navigate("https://bestvoip.forumotion.com/post?f=3&mode=newtopic")
        voiptraffic.Navigate("https://voiptraffic.forumotion.com/post?f=7&mode=newtopic")
        reload()
        'MsgBox(ComboBoxSubject.Items.Count)
    End Sub

    Private Sub Label2_Click(sender As Object, e As EventArgs) Handles Label2.Click
        Clipboard.SetText(Label2.Text)
    End Sub

    Private Sub Label3_Click(sender As Object, e As EventArgs) Handles Label3.Click
        Clipboard.SetText(Label3.Text)
    End Sub

    Private Sub LabelUser2_Click(sender As Object, e As EventArgs) Handles LabelUser2.Click
        Clipboard.SetText(LabelUser2.Text)
    End Sub

    Private Sub LabelPass2_Click(sender As Object, e As EventArgs) Handles LabelPass2.Click
        Clipboard.SetText(LabelPass2.Text)
    End Sub

    Private Sub lblSubject_Click(sender As Object, e As EventArgs) Handles lblSubject.Click
        Clipboard.SetText("We are looking for Direct supplier")
    End Sub

    Private Sub lblBody_Click(sender As Object, e As EventArgs) Handles lblBody.Click
        Clipboard.SetText("Dear Provider,

Greetings from Filasco America LLC

We're looking Direct provider for A-Z NCLI, CLI and pure TDM terminations.
Thanks.

Best regards
Jacob Hoffman
Sales Executive Officer at Filasco America LLC.
Skype: live:jacob.hoffman_8
Email :jacob.hoffman@filasco.net
Website: http://www.filasco.net
Phone: +1(646) 691-2952
Fax: +1(866) 996-8535
Corporate Office: 47th street, Suite B3, Brooklyn, NY 11220")
    End Sub

    Private Sub lblForumVoipUser_Click(sender As Object, e As EventArgs) Handles lblForumVoipUser.Click
        Clipboard.SetText(lblForumVoipUser.Text)
    End Sub

    Private Sub lblForumVoipPass_Click(sender As Object, e As EventArgs) Handles lblForumVoipPass.Click
        Clipboard.SetText(lblForumVoipPass.Text)
    End Sub

    Private Sub lblVoipHelpUser_Click(sender As Object, e As EventArgs) Handles lblVoipHelpUser.Click
        Clipboard.SetText(lblVoipHelpUser.Text)
    End Sub

    Private Sub lblVoipHelpPass_Click(sender As Object, e As EventArgs) Handles lblVoipHelpPass.Click
        Clipboard.SetText(lblVoipHelpPass.Text)
    End Sub

    Private Sub lblBestVoipUser_Click(sender As Object, e As EventArgs) Handles lblBestVoipUser.Click
        Clipboard.SetText(lblBestVoipUser.Text)
    End Sub

    Private Sub lblBestVoipPass_Click(sender As Object, e As EventArgs) Handles lblBestVoipPass.Click
        Clipboard.SetText(lblBestVoipPass.Text)
    End Sub

    Private Sub lblVoipTrafficUser_Click(sender As Object, e As EventArgs) Handles lblVoipTrafficUser.Click
        Clipboard.SetText(lblVoipTrafficUser.Text)
    End Sub

    Private Sub lblVoipTrafficPass_Click(sender As Object, e As EventArgs) Handles lblVoipTrafficPass.Click
        Clipboard.SetText(lblVoipTrafficPass.Text)
    End Sub

    Private Sub Button1_Click_3(sender As Object, e As EventArgs) Handles btnSaveAds.Click
        If MsgBox("Are you sure to insert?", MsgBoxStyle.YesNo, Title:="Notice") = vbYes Then
            mycmd.Connection = myconnection.open
            mycmd.CommandText = "insert into attempt_to_registration (User_Name,Comments)
            values('" & txtSubject.Text & "','" & txtBody.Text & "')"
            mycmd.ExecuteNonQuery()
            myconnection.close()
            MsgBox("Saved!")
        End If
    End Sub

    Private Sub btnPost_Click(sender As Object, e As EventArgs) Handles btnPost.Click
        Try
            callsMarket.Document.GetElementById("post").InvokeMember("click")
            calltermination.Document.GetElementById("post").InvokeMember("click")

            Dim inputTagCollection = forumVoip.Document.GetElementsByTagName("input")
            For Each inputTag As HtmlElement In inputTagCollection
                If inputTag.OuterHtml.Contains("post") Then
                    inputTag.InvokeMember("click")
                End If
            Next
            Dim inputTagCollection1 = voiphelp.Document.GetElementsByTagName("input")
            For Each inputTag As HtmlElement In inputTagCollection1
                If inputTag.OuterHtml.Contains("post") Then
                    inputTag.InvokeMember("click")
                End If
            Next
            Dim inputTagCollection2 = bestVoip.Document.GetElementsByTagName("input")
            For Each inputTag As HtmlElement In inputTagCollection2
                If inputTag.OuterHtml.Contains("post") Then
                    inputTag.InvokeMember("click")
                End If
            Next
            Dim inputTagCollection3 = voiptraffic.Document.GetElementsByTagName("input")
            For Each inputTag As HtmlElement In inputTagCollection3
                If inputTag.OuterHtml.Contains("post") Then
                    inputTag.InvokeMember("click")
                End If
            Next
        Catch ex As Exception

        End Try
    End Sub

    Private Sub btnUpdate_Click(sender As Object, e As EventArgs) Handles btnUpdate.Click
        If MsgBox("Are you sure to update?", MsgBoxStyle.YesNo, Title:="Notice") = vbYes Then
            mycmd.Connection = myconnection.open
            mycmd.CommandText = "update attempt_to_registration set User_Name='" + txtSubject.Text + "', Comments='" + txtBody.Text + "' where User_Name='" + ComboBoxSubject.Text + "'"
            objreader = mycmd.ExecuteReader
            myconnection.close()
            MsgBox("Success!")
        End If
    End Sub

    Private Sub btnDelete_Click(sender As Object, e As EventArgs) Handles btnDelete.Click
        If MsgBox("Are you sure to delete?", MsgBoxStyle.YesNo, Title:="Notice") = vbYes Then
            mycmd.Connection = myconnection.open
            mycmd.CommandText = "delete from attempt_to_registration where User_Name =  '" + ComboBoxSubject.Text + "'"
            objreader = mycmd.ExecuteReader
            myconnection.close()
            MsgBox("Success!")
        End If
    End Sub
    Private Sub reload()
        ComboBoxSubject.Items.Clear()
        Try
            Dim query As String
            Mysqlconn = myconnection.open
            query = "select * from attempt_to_registration"
            mycmd = New MySqlCommand(query, Mysqlconn)
            objreader = mycmd.ExecuteReader

            While objreader.Read
                Dim sName = objreader.GetString("User_Name")

                ComboBoxSubject.Items.Add(sName)
            End While
            myconnection.close()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub btnStop_Click(sender As Object, e As EventArgs) Handles btnStop.Click
        Timer1.Dispose()
    End Sub

    Private Sub btnSwitch_Click(sender As Object, e As EventArgs) Handles btnSwitch.Click
        Form1.Show()
        Me.Close()
    End Sub

End Class
