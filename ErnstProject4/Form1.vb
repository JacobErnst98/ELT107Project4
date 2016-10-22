Public Class Form1
    Dim A, B, C, D, Min, Max, Act As Decimal

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ListBox4.SelectedIndex = 1 ' set gold as defualt
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click 'when button clicked
        If ListBox1.SelectedIndex = 0 Then
            MsgBox("Please select a color other than Black for the first band") ' error check
            Exit Sub
        End If
        If (ListBox1.SelectedIndex Or ListBox2.SelectedIndex Or ListBox3.SelectedIndex Or ListBox4.SelectedIndex) = -1 Then
            MsgBox("Please select a color form each box")
            Exit Sub
        End If
        A = ListBox1.SelectedIndex 'set vars
        B = ListBox2.SelectedIndex
        C = 10 ^ ListBox3.SelectedIndex
        If ListBox4.SelectedIndex = 0 Then D = 0.02 'tolerance percentages
        If ListBox4.SelectedIndex = 1 Then D = 0.05
        ' MsgBox(A & " " & B & " " & C) debug
        Label1.BackColor = Color.FromName(ListBox1.SelectedItem) 'band colors
        Label2.BackColor = Color.FromName(ListBox2.SelectedItem)
        Label3.BackColor = Color.FromName(ListBox3.SelectedItem)
        Label4.BackColor = Color.FromName(ListBox4.SelectedItem)
        Label1.Text = "                    " 'band sizes
        Label2.Text = "                    "
        Label3.Text = "                    "
        Label4.Text = "                    "
        Act = ((A & B) * C) 'calculate values
        Label11.Text = "Ractual " & Act
        Min = ((A & B) * C) - ((A & B) * C * D)
        Max = ((A & B) * C) + ((A & B) * C * D)
        Label5.Text = "Min " & Min & " 0hms"
        Label6.Text = "Max " & Max & " 0hms"
        If Min > 1000 Then   'calculate units 
            Min = Min / 1000
            Label5.Text = ("Min " & Min & " Kohms")
        End If
        If Min > 1000000 Then
            Min = Min / 1000000
            Label5.Text = ("Min " & Min & " Mohms")
        End If
        If Max > 1000 Then
            Max = Max / 1000
            Label6.Text = ("Max " & Max & " Kohms")
        End If
        If Max > 1000000 Then
            Max = Max / 1000000
            Label6.Text = ("Max s" & Max & " Mohms")
        End If

    End Sub
End Class
