Public Class Form1
    Private Sub Button1_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load


        Form2.Show()
        Me.Hide()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs)


    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        Dim x As Integer = ProgressBar1.Value


        Dim read As String = SerialPort1.ReadExisting
        'Label1.Text = read
        'If read.Contains("B") Then
        'x += 1
        'If x < 5 Then
        'ProgressBar1.Value = x
        'End If
        'ElseIf read.Contains("Y") Then
        'x += 1
        'If x < 5 Then
        'ProgressBar1.Value = x
        'End If
        'ElseIf read.Contains("X") Then
        'x += 1
        'If x < 5 Then
        'ProgressBar1.Value = x
        'End If
        ' If



        If read.Contains("1") Then
            For x = 1 To 100
                ProgressBar1.Value += 1
                Threading.Thread.Sleep(30)
            Next
            Form2.Show()
            Me.Hide()
        ElseIf read.Contains("0") Then
            MsgBox("Incorrect")
        End If




    End Sub

    Private Sub PictureBox2_Click(sender As Object, e As EventArgs) Handles PictureBox2.Click
        Application.Exit()
    End Sub

    Private Sub PictureBox3_Click(sender As Object, e As EventArgs) Handles PictureBox3.Click
        Me.WindowState = FormWindowState.Minimized
    End Sub

    Private Sub Button1_Click_1(sender As Object, e As EventArgs) Handles Button1.Click
        CmbPort.Items.Clear()
        Dim myPort As Array
        Dim i As Integer
        myPort = IO.Ports.SerialPort.GetPortNames()
        CmbPort.Items.AddRange(myPort)
        i = CmbPort.Items.Count
        i = i - i
        Try
            CmbPort.SelectedIndex = i
        Catch ex As Exception
            Dim result As DialogResult
            result = MessageBox.Show("com port not detected", "Warning !!!", MessageBoxButtons.OK)
            CmbPort.Text = ""
            CmbPort.Items.Clear()
            Call Form1_Load(Me, e)
        End Try
        'BtnCon.Enabled = True
        'BtnCon.BringToFront()
        CmbPort.DroppedDown = True
    End Sub

    Private Sub Button2_Click_1(sender As Object, e As EventArgs) Handles Button2.Click


        SerialPort1.BaudRate = 9600
        If CmbPort.Text.Contains("COM") Then
            SerialPort1.PortName = CmbPort.SelectedItem
            Try
                SerialPort1.Open()
                Timer1.Start()
                Label1.Visible = True
                Label2.Visible = False
            Catch ex As Exception
                MsgBox("Serial Port busy or not available.")
            End Try
        Else
            MsgBox("Select a port")
        End If





    End Sub

    Private Sub Label1_Click(sender As Object, e As EventArgs) Handles Label1.Click

    End Sub
End Class
