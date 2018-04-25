Imports SpeechLib

Public Class Form1
    Dim s As New SpVoice

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        On Error Resume Next
        s.Speak(RichTextBox1.Text)
        TLabel2.Text = "Statue : Ready..."
    End Sub

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        If ComboBox1.Text <> ComboBox1.SelectedItem Or ComboBox1.Text = "" Then MsgBox("Please choose the font", MsgBoxStyle.Exclamation)
        If ComboBox2.Text <> ComboBox2.SelectedItem Or ComboBox2.Text = "" Then MsgBox("Please choose the style", MsgBoxStyle.Exclamation)
        If ComboBox3.Text <> ComboBox3.SelectedItem Or ComboBox3.Text = "" Then MsgBox("Please choose the color", MsgBoxStyle.Exclamation)
        If ComboBox1.Text = ComboBox1.SelectedItem And ComboBox2.Text = ComboBox2.SelectedItem And ComboBox3.Text = ComboBox3.SelectedItem Then
            Dim f As FontStyle = FontStyle.Regular
            Dim c As Color = Color.Black
            Select Case ComboBox2.SelectedIndex
                Case 0
                    f = FontStyle.Regular
                Case 1
                    f = FontStyle.Bold
                Case 2
                    f = FontStyle.Italic
                Case 3
                    f = FontStyle.Bold + FontStyle.Italic
                Case 4
                    f = FontStyle.Underline
                Case 5
                    f = FontStyle.Bold + FontStyle.Underline
                Case 6
                    f = FontStyle.Italic + FontStyle.Underline
            End Select
            Select Case ComboBox3.SelectedIndex
                Case 0
                    c = Color.Black
                Case 1
                    c = Color.White
                Case 2
                    c = Color.Yellow
                Case 3
                    c = Color.Orange
                Case 4
                    c = Color.Red
                Case 5
                    c = Color.Green
                Case 6
                    c = Color.Blue
                Case 7
                    c = Color.Gray
            End Select
            RichTextBox1.Font = New Drawing.Font(ComboBox1.Text, NumericUpDown1.Value, f, GraphicsUnit.Point)
            RichTextBox1.ForeColor = c
        End If
    End Sub

    Private Sub FontDialog1_Apply(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FontDialog1.Apply
        RichTextBox1.Font = FontDialog1.Font
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        If FontDialog1.ShowDialog = DialogResult.OK Then RichTextBox1.Font = FontDialog1.Font
    End Sub

    Private Sub TDimer_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TDimer.Tick
        TLabel1.Text = "Time : " & TimeOfDay & " , Date : " & Today & " ,"
        TLabel3.Text = ListBox1.SelectedItem
        If ListBox1.SelectedIndex >= 13 Then
            ListBox1.SelectedIndex = 0
        Else
            ListBox1.SelectedIndex = ListBox1.SelectedIndex + +1
        End If
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        RichTextBox1.Text = ""
        s.Pause()
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Me.Close()
    End Sub

    Private Sub Button1_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Button1.MouseDown
        TLabel2.Text = "Statue : Speaking"
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        OpenFileDialog1.ShowDialog()
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        SaveFileDialog1.ShowDialog()
    End Sub

    Private Sub OpenFileDialog1_FileOk(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles OpenFileDialog1.FileOk
        On Error Resume Next
        RichTextBox1.Text = RichTextBox1.Text & My.Computer.FileSystem.ReadAllText(OpenFileDialog1.FileName)
    End Sub

    Private Sub SaveFileDialog1_FileOk(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles SaveFileDialog1.FileOk
        My.Computer.FileSystem.WriteAllText(SaveFileDialog1.FileName, RichTextBox1.Text, True)
    End Sub

    Private Sub NewToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NewToolStripMenuItem.Click
        RichTextBox1.Text = ""
        s.Pause()
    End Sub

    Private Sub OpenToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OpenToolStripMenuItem.Click
        OpenFileDialog1.ShowDialog()
    End Sub

    Private Sub SaveToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SaveToolStripMenuItem.Click
        SaveFileDialog1.ShowDialog()
    End Sub

    Private Sub ExitToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExitToolStripMenuItem.Click
        Me.Close()
    End Sub

    Private Sub UndoToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UndoToolStripMenuItem.Click
        RichTextBox1.Undo()
    End Sub

    Private Sub RedoToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RedoToolStripMenuItem.Click
        RichTextBox1.Redo()
    End Sub

    Private Sub CutToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CutToolStripMenuItem.Click
        RichTextBox1.Cut()
    End Sub

    Private Sub CopyToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CopyToolStripMenuItem.Click
        RichTextBox1.Copy()
    End Sub

    Private Sub PasteToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PasteToolStripMenuItem.Click
        RichTextBox1.Paste()
    End Sub

    Private Sub SelectAllToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SelectAllToolStripMenuItem.Click
        RichTextBox1.SelectAll()
    End Sub

    Private Sub FontToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FontToolStripMenuItem.Click
        If FontDialog1.ShowDialog = DialogResult.OK Then RichTextBox1.Font = FontDialog1.Font
    End Sub

    Private Sub FormColorToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FormColorToolStripMenuItem.Click
        If ColorDialog1.ShowDialog = DialogResult.OK Then Me.BackColor = ColorDialog1.Color
    End Sub

    Private Sub MenuColorToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuColorToolStripMenuItem.Click
        If ColorDialog1.ShowDialog = DialogResult.OK Then MenuStrip1.BackColor = ColorDialog1.Color
    End Sub

    Private Sub TextColorToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextColorToolStripMenuItem.Click
        If ColorDialog1.ShowDialog = DialogResult.OK Then RichTextBox1.ForeColor = ColorDialog1.Color
    End Sub

    Private Sub ResetToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ResetToolStripMenuItem.Click
        MenuStrip1.BackColor = Color.WhiteSmoke
        Me.BackColor = Color.WhiteSmoke
        RichTextBox1.ForeColor = Color.Black
        RichTextBox1.Font = New Drawing.Font("Arial", 12, FontStyle.Regular, GraphicsUnit.Point)
    End Sub

    Private Sub AboutToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AboutToolStripMenuItem.Click
        MsgBox("Text Reader En Version", MsgBoxStyle.Information, "About")
    End Sub

    Private Sub TipToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TipToolStripMenuItem.Click
        MsgBox("The Program'll paused while Reading")
    End Sub

    Private Sub CreditsToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CreditsToolStripMenuItem.Click
        MsgBox("All rights are reserved to Mohamed Selim ©  2014", MsgBoxStyle.Information, "Credits")
    End Sub

    Private Sub CutToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CutToolStripMenuItem1.Click
        RichTextBox1.Cut()
    End Sub

    Private Sub CopyToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CopyToolStripMenuItem1.Click
        RichTextBox1.Copy()
    End Sub

    Private Sub PasteToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PasteToolStripMenuItem1.Click
        RichTextBox1.Paste()
    End Sub

    Private Sub TextColorToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextColorToolStripMenuItem1.Click
        If ColorDialog1.ShowDialog = DialogResult.OK Then RichTextBox1.ForeColor = ColorDialog1.Color
    End Sub

    Private Sub FontToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FontToolStripMenuItem1.Click
        If FontDialog1.ShowDialog = DialogResult.OK Then RichTextBox1.Font = FontDialog1.Font
    End Sub

    Private Sub CreditsToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CreditsToolStripMenuItem1.Click
        MsgBox("All rights are reserved to Mohamed Selim ©  2014", MsgBoxStyle.Information, "Credits")
    End Sub

    Private Sub RichTextBox1_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles RichTextBox1.MouseUp
        Dim c As Control = CType(sender, Control)
        If e.Button = Windows.Forms.MouseButtons.Right Then
            Me.ContextMenuStrip1.Show(c, e.Location, ToolStripDropDownDirection.Default)
        End If
    End Sub

    Private Sub Form1_SizeChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.SizeChanged
        GroupBox1.Location = New Point(Me.Size.Width / 2 - GroupBox1.Size.Width / 2 - 8, 27)
        RichTextBox1.Size = New Size(Me.Size.Width - 40, Me.Size.Height - 300)
        GroupBox2.Location = New Point(Me.Size.Width / 2 - GroupBox2.Size.Width / 2 - 8, RichTextBox1.Location.Y + RichTextBox1.Size.Height + 6)
    End Sub
End Class
