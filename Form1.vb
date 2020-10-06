Public Class Form1
    Public deg2rad 'As Double  'Decimal
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        On Error Resume Next
        TextBox13.Clear()
        TextBox14.Clear()
        TextBox15.Clear()
        TextBox16.Clear()
        TextBox17.Clear()

        Label30.Text = ""
        Label31.Text = ""
        Label32.Text = ""
        Label33.Text = ""

        Label44.Text = ""
        Label45.Text = ""
        Label43.Text = ""
        Label42.Text = ""

    End Sub

    Private Sub PictureBox3_Click(sender As Object, e As EventArgs) Handles PictureBox3.Click
        On Error Resume Next
        deg2rad = Math.PI / 180

        TextBox13.Text = Format(TextBox1.Text + (TextBox2.Text / 60) + (TextBox3.Text / 3600), "#0.###0")
        TextBox14.Text = Format(TextBox4.Text + (TextBox5.Text / 60) + (TextBox6.Text / 3600), "#0.###0")
        TextBox15.Text = Format(TextBox7.Text + (TextBox8.Text / 60) + (TextBox9.Text / 3600), "#0.###0")
        TextBox16.Text = Format(TextBox10.Text + (TextBox11.Text / 60) + (TextBox12.Text / 3600), "#0.###0")

        TextBox17.Text = Format(Math.Acos(Math.Cos((90 - TextBox13.Text) * deg2rad) * Math.Cos((90 - TextBox15.Text) * deg2rad) + Math.Sin((90 - TextBox13.Text) * deg2rad) * Math.Sin((90 - TextBox15.Text) * deg2rad) * Math.Cos((TextBox14.Text - TextBox16.Text) * deg2rad)) * 6378.137, "#,000.###0")

        If TextBox13.Text >= 0 Then
            Label30.Text = "N"
        ElseIf TextBox13.Text <= 0 Then
            Label30.Text = "S"
        End If

        If TextBox14.Text >= 0 Then
            Label31.Text = "E"
        ElseIf TextBox14.Text <= 0 Then
            Label31.Text = "W"
        End If


        If TextBox15.Text >= 0 Then
            Label32.Text = "N"
        ElseIf TextBox15.Text <= 0 Then
            Label32.Text = "S"
        End If

        If TextBox16.Text >= 0 Then
            Label33.Text = "E"
        ElseIf TextBox16.Text <= 0 Then
            Label33.Text = "W"
        End If

        Label45.Text = Label30.Text
        Label44.Text = Label31.Text
        Label43.Text = Label32.Text
        Label42.Text = Label33.Text


    End Sub

    Private Sub PictureBox2_Click(sender As Object, e As EventArgs) Handles PictureBox2.Click
        On Error Resume Next
        deg2rad = Math.PI / 180

        TextBox20.Text = Int(TextBox18.Text)
        TextBox21.Text = Int((TextBox18.Text - TextBox20.Text) * 60)
        TextBox22.Text = Format((((TextBox18.Text - TextBox20.Text) * 60) - Int((TextBox18.Text - TextBox20.Text) * 60)) * 60, "#0.###0")
        TextBox23.Text = Int(TextBox19.Text)
        TextBox24.Text = Int((TextBox19.Text - TextBox23.Text) * 60)
        TextBox25.Text = Format((((TextBox19.Text - TextBox23.Text) * 60) - Int((TextBox19.Text - TextBox23.Text) * 60)) * 60, "#0.###0")

        TextBox28.Text = Int(TextBox26.Text)
        TextBox29.Text = Int((TextBox26.Text - TextBox28.Text) * 60)
        TextBox30.Text = Format((((TextBox26.Text - TextBox28.Text) * 60) - Int((TextBox26.Text - TextBox28.Text) * 60)) * 60, "#0.###0")
        TextBox31.Text = Int(TextBox27.Text)
        TextBox32.Text = Int((TextBox27.Text - TextBox31.Text) * 60)
        TextBox33.Text = Format((((TextBox27.Text - TextBox31.Text) * 60) - Int((TextBox27.Text - TextBox31.Text) * 60)) * 60, "#0.###0")

        TextBox34.Text = Format(Math.Acos(Math.Cos((90 - TextBox18.Text) * deg2rad) * Math.Cos((90 - TextBox26.Text) * deg2rad) + Math.Sin((90 - TextBox18.Text) * deg2rad) * Math.Sin((90 - TextBox26.Text) * deg2rad) * Math.Cos((TextBox19.Text - TextBox27.Text) * deg2rad)) * 6378.137, "#,000.###0")

        If TextBox18.Text >= 0 Then
            Label34.Text = "N"
            Label39.Text = "N"
        ElseIf TextBox18.Text <= 0 Then
            Label34.Text = "S"
            Label39.Text = "S"
        End If

        If TextBox19.Text >= 0 Then
            Label35.Text = "E"
            Label38.Text = "E"
        ElseIf TextBox19.Text <= 0 Then
            Label35.Text = "W"
            Label38.Text = "W"
        End If




        If TextBox26.Text >= 0 Then
            Label37.Text = "N"
            Label41.Text = "N"
        ElseIf TextBox26.Text <= 0 Then
            Label37.Text = "S"
            Label41.Text = "S"
        End If

        If TextBox27.Text >= 0 Then
            Label36.Text = "E"
            Label40.Text = "E"
        ElseIf TextBox27.Text <= 0 Then
            Label36.Text = "W"
            Label40.Text = "W"
        End If




    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        On Error Resume Next

        TextBox20.Clear()
        TextBox21.Clear()
        TextBox22.Clear()
        TextBox23.Clear()
        TextBox24.Clear()
        TextBox25.Clear()
        TextBox28.Clear()
        TextBox29.Clear()
        TextBox30.Clear()
        TextBox31.Clear()
        TextBox32.Clear()
        TextBox33.Clear()
        TextBox34.Clear()

        Label34.Text = ""
        Label35.Text = ""
        Label38.Text = ""
        Label39.Text = ""
        Label37.Text = ""
        Label36.Text = ""
        Label41.Text = ""
        Label40.Text = ""

    End Sub

    Private Sub Label46_Click(sender As Object, e As EventArgs) Handles Label46.Click
        System.Diagnostics.Process.Start("chrome.exe", "http://antenna-design.tistory.com/27")
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        On Error Resume Next
        Dim sOutput As String
        Dim lvItem As ListViewItem
        'Clipboard.SetDataObject(TextBox13.Text & Label30.Text & " " & TextBox14.Text & Label31.Text)
        Clipboard.SetDataObject(TextBox13.Text & " " & TextBox14.Text)
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        On Error Resume Next
        Dim sOutput As String
        Dim lvItem As ListViewItem
        'Clipboard.SetDataObject(TextBox15.Text & Label32.Text & " " & TextBox16.Text & Label33.Text)
        Clipboard.SetDataObject(TextBox15.Text & " " & TextBox16.Text)
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        On Error Resume Next
        Dim sOutput As String
        Dim lvItem As ListViewItem
        Clipboard.SetDataObject(TextBox20.Text & "°" & TextBox21.Text & "'" & TextBox22.Text & "””" & Label39.Text & " " & TextBox23.Text & "°" & TextBox24.Text & "'" & TextBox25.Text & "””" & Label38.Text)

    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        On Error Resume Next
        Dim sOutput As String
        Dim lvItem As ListViewItem
        Clipboard.SetDataObject(TextBox28.Text & "°" & TextBox29.Text & "'" & TextBox30.Text & "””" & Label41.Text & " " & TextBox31.Text & "°" & TextBox32.Text & "'" & TextBox33.Text & "””" & Label40.Text)

    End Sub

    Private Sub Label51_Click(sender As Object, e As EventArgs) Handles Label51.Click
        System.Diagnostics.Process.Start("chrome.exe", "https://maps.google.co.kr/maps?hl=ko&tab=rl")
    End Sub

    Private Sub Label52_Click(sender As Object, e As EventArgs) Handles Label52.Click

        System.Diagnostics.Process.Start("iexplore.exe", "https://maps.google.co.kr/maps?hl=ko&tab=rl")
    End Sub

End Class
