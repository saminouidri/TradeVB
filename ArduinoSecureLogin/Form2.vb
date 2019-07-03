Public Class Form2
    Dim AAPL As Integer = 195
    Dim BAYN As Integer = 60
    Dim FB As Integer = 175
    Dim GOOGL As Integer = 1071
    Dim MSFT As Integer = 130
    Dim AMZN As Integer = 1800
    Dim NOVN As Integer = 87
    Dim UBS As Integer = 12
    Dim MaxTab As Integer = 0
    Dim MaxTab2 As Integer = 0
    Dim MaxTab3 As Integer = 0
    Dim MaxTab4 As Integer = 0
    Dim MaxTab5 As Integer = 0
    Dim MaxTab6 As Integer = 0
    Dim MaxTab7 As Integer = 0
    Dim MaxTab8 As Integer = 0
    Dim counter As Integer = 0

    Dim Dollar As Double = 1
    Dim flip As Integer = 0
    Dim check As Integer = 0

    Dim balance As Integer = 100
    Private Sub PictureBox2_Click(sender As Object, e As EventArgs) Handles PictureBox2.Click
        Application.Exit()
    End Sub

    Private Sub PictureBox3_Click(sender As Object, e As EventArgs) Handles PictureBox3.Click
        Me.WindowState = FormWindowState.Minimized
    End Sub
    Private StockTab As TabPage


    Public Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim NewLB As Label = New Label With {.Text = ComboBox1.Text}
        NewLB.Size = New System.Drawing.Size(159, 23)
        NewLB.Location = New System.Drawing.Point(12, 10)
        NewLB.Font = New Font("Bebas Neue", 18, FontStyle.Regular)

        Dim PriceLB As Label = New Label
        PriceLB.Size = New System.Drawing.Size(159, 23)
        PriceLB.Location = New System.Drawing.Point(550, 10)
        PriceLB.Font = New Font("Bebas Neue", 18, FontStyle.Regular)
        PriceLB.ForeColor = System.Drawing.Color.Green

        If ComboBox1.Text.Contains("AAPL") Then
            If MaxTab < 1 Then
                PriceAAPL.Visible = True
                PriceAAPL.Text = AAPL
                PriceAAPL.Size = New System.Drawing.Size(159, 23)
                PriceAAPL.Location = New System.Drawing.Point(470, 10)
                PriceAAPL.Font = New Font("Bebas Neue", 18, FontStyle.Regular)
                PriceAAPL.ForeColor = System.Drawing.Color.Green
                StockTab = New TabPage()
                StockTab.Text = ComboBox1.Text
                TabControl1.TabPages.Add(StockTab)
                StockTab.Controls.Add(PriceAAPL)
                StockTab.Controls.Add(NewLB)
                ChartAAPL.Visible = True
                StockTab.Controls.Add(ChartAAPL)
                MaxTab += 1
            End If

        ElseIf ComboBox1.Text.Contains("BAYN") Then
            If MaxTab2 < 1 Then
                PriceBAYN.Visible = True
                PriceBAYN.Text = BAYN
                PriceBAYN.Size = New System.Drawing.Size(159, 23)
                PriceBAYN.Location = New System.Drawing.Point(470, 10)
                PriceBAYN.Font = New Font("Bebas Neue", 18, FontStyle.Regular)
                PriceBAYN.ForeColor = System.Drawing.Color.Green
                StockTab = New TabPage()
                StockTab.Text = ComboBox1.Text
                TabControl1.TabPages.Add(StockTab)
                StockTab.Controls.Add(PriceBAYN)
                StockTab.Controls.Add(NewLB)
                ChartBAYN.Visible = True
                StockTab.Controls.Add(ChartBAYN)
                MaxTab2 += 1
            End If
        ElseIf ComboBox1.Text.Contains("FB") Then
            If MaxTab3 < 1 Then
                PriceFB.Visible = True
                PriceFB.Text = FB
                PriceFB.Size = New System.Drawing.Size(159, 23)
                PriceFB.Location = New System.Drawing.Point(470, 10)
                PriceFB.Font = New Font("Bebas Neue", 18, FontStyle.Regular)
                PriceFB.ForeColor = System.Drawing.Color.Green
                StockTab = New TabPage()
                StockTab.Text = ComboBox1.Text
                TabControl1.TabPages.Add(StockTab)
                StockTab.Controls.Add(PriceFB)
                StockTab.Controls.Add(NewLB)
                ChartFB.Visible = True
                StockTab.Controls.Add(ChartFB)
                MaxTab3 += 1
            End If
        ElseIf ComboBox1.Text.Contains("GOOGL") Then
            If MaxTab4 < 1 Then
                PriceGOOGL.Visible = True
                PriceGOOGL.Text = GOOGL
                PriceGOOGL.Size = New System.Drawing.Size(159, 23)
                PriceGOOGL.Location = New System.Drawing.Point(470, 10)
                PriceGOOGL.Font = New Font("Bebas Neue", 18, FontStyle.Regular)
                PriceGOOGL.ForeColor = System.Drawing.Color.Green
                StockTab = New TabPage()
                StockTab.Text = ComboBox1.Text
                TabControl1.TabPages.Add(StockTab)
                StockTab.Controls.Add(PriceGOOGL)
                StockTab.Controls.Add(NewLB)
                ChartGOOGL.Visible = True
                StockTab.Controls.Add(ChartGOOGL)
                MaxTab4 += 1
            End If
        ElseIf ComboBox1.Text.Contains("MSFT") Then
            If MaxTab5 < 1 Then
                PriceMSFT.Visible = True
                PriceMSFT.Text = MSFT
                PriceMSFT.Size = New System.Drawing.Size(159, 23)
                PriceMSFT.Location = New System.Drawing.Point(470, 10)
                PriceMSFT.Font = New Font("Bebas Neue", 18, FontStyle.Regular)
                PriceMSFT.ForeColor = System.Drawing.Color.Green
                StockTab = New TabPage()
                StockTab.Text = ComboBox1.Text
                TabControl1.TabPages.Add(StockTab)
                StockTab.Controls.Add(PriceMSFT)
                StockTab.Controls.Add(NewLB)
                ChartMSFT.Visible = True
                StockTab.Controls.Add(ChartMSFT)
                MaxTab5 += 1
            End If
        ElseIf ComboBox1.Text.Contains("AMZN") Then
            If MaxTab6 < 1 Then
                PriceAMZN.Visible = True
                PriceAMZN.Text = AMZN
                PriceAMZN.Size = New System.Drawing.Size(159, 23)
                PriceAMZN.Location = New System.Drawing.Point(470, 10)
                PriceAMZN.Font = New Font("Bebas Neue", 18, FontStyle.Regular)
                PriceAMZN.ForeColor = System.Drawing.Color.Green
                StockTab = New TabPage()
                StockTab.Text = ComboBox1.Text
                TabControl1.TabPages.Add(StockTab)
                StockTab.Controls.Add(PriceAMZN)
                StockTab.Controls.Add(NewLB)
                ChartAMZN.Visible = True
                StockTab.Controls.Add(ChartAMZN)
                MaxTab6 += 1
            End If
        ElseIf ComboBox1.Text.Contains("NOVN") Then
            If MaxTab7 < 1 Then
                PriceNOVN.Visible = True
                PriceNOVN.Text = NOVN
                PriceNOVN.Size = New System.Drawing.Size(159, 23)
                PriceNOVN.Location = New System.Drawing.Point(470, 10)
                PriceNOVN.Font = New Font("Bebas Neue", 18, FontStyle.Regular)
                PriceNOVN.ForeColor = System.Drawing.Color.Green
                StockTab = New TabPage()
                StockTab.Text = ComboBox1.Text
                TabControl1.TabPages.Add(StockTab)
                StockTab.Controls.Add(PriceNOVN)
                StockTab.Controls.Add(NewLB)
                ChartNOVN.Visible = True
                StockTab.Controls.Add(ChartNOVN)
                MaxTab7 += 1
            End If
        ElseIf ComboBox1.Text.Contains("UBS") Then
            If MaxTab8 < 1 Then
                PriceUBS.Visible = True
                PriceUBS.Text = UBS
                PriceUBS.Size = New System.Drawing.Size(159, 23)
                PriceUBS.Location = New System.Drawing.Point(470, 10)
                PriceUBS.Font = New Font("Bebas Neue", 18, FontStyle.Regular)
                PriceUBS.ForeColor = System.Drawing.Color.Green
                StockTab = New TabPage()
                StockTab.Text = ComboBox1.Text
                TabControl1.TabPages.Add(StockTab)
                StockTab.Controls.Add(PriceUBS)
                StockTab.Controls.Add(NewLB)
                ChartUBS.Visible = True
                StockTab.Controls.Add(ChartUBS)
                MaxTab8 += 1
            End If
        End If


        ' StockTab = New TabPage()
        ' StockTab.Text = ComboBox1.Text
        ' TabControl1.TabPages.Add(StockTab)

        ' StockTab.Controls.Add(PriceLB)
        ' StockTab.Controls.Add(NewLB)


    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        Dim Variation As Integer = CInt((Rnd() * ((-1) ^ Int(Rnd() * 10))) * 20)
        Dim Variation1 As Integer = CInt((Rnd() * ((-1) ^ Int(Rnd() * 10))) * 20)
        Dim Variation2 As Integer = CInt((Rnd() * ((-1) ^ Int(Rnd() * 10))) * 20)
        Dim Variation3 As Integer = CInt((Rnd() * ((-1) ^ Int(Rnd() * 10))) * 20)
        Dim Variation4 As Integer = CInt((Rnd() * ((-1) ^ Int(Rnd() * 10))) * 20)
        Dim Variation5 As Integer = CInt((Rnd() * ((-1) ^ Int(Rnd() * 10))) * 20)
        Dim Variation6 As Integer = CInt((Rnd() * ((-1) ^ Int(Rnd() * 10))) * 20)
        Dim Variation7 As Integer = CInt((Rnd() * ((-1) ^ Int(Rnd() * 10))) * 20)
        'Dim Variation8 As Double = CInt((Rnd() * ((-1) ^ Int(Rnd() * 0))) * 1)
        'Dim Change As String

        Dim r As New Random() 'Should be declared at the topmost level
        Dim value As Double = r.NextDouble()
        value = value * 0.1
        flip += 1
        check = flip Mod 2
        If check = 0 Then
            value = value * 1
        Else
            value = value * -1
        End If
        Dollar += value
        'Dollar -= 0.3
        Chart1.ChartAreas(0).AxisX.Interval = 1
        Chart1.ChartAreas(0).AxisX.Minimum = 0
        Chart1.ChartAreas(0).AxisX.LabelStyle.Angle = 90
        Chart1.ChartAreas(0).AxisY.Interval = 0.25
        Chart1.ChartAreas(0).AxisY.Minimum = 0
        Chart1.Series(0).Points.AddXY(counter, Dollar)

        AAPL += Variation
        ChartAAPL.ChartAreas(0).AxisX.Interval = 1
        ChartAAPL.ChartAreas(0).AxisX.Minimum = 0
        ChartAAPL.ChartAreas(0).AxisX.LabelStyle.Angle = 90
        ChartAAPL.ChartAreas(0).AxisY.Interval = 50
        ChartAAPL.ChartAreas(0).AxisY.Minimum = 0
        ChartAAPL.Series(0).Points.AddXY(counter, AAPL)
        If Variation < 0 Then
            PriceAAPL.ForeColor = System.Drawing.Color.Red
            PriceAAPL.Text = "(" & Variation & "$) " & AAPL
        Else
            PriceAAPL.ForeColor = System.Drawing.Color.Green
            PriceAAPL.Text = "(+" & Variation & "$) " & AAPL
        End If
        'PriceAAPL.Text = "(" & Variation & "$) " & AAPL
        'MsgBox(Variation)
        BAYN += Variation1
        ChartBAYN.ChartAreas(0).AxisX.Interval = 1
        ChartBAYN.ChartAreas(0).AxisX.Minimum = 0
        ChartBAYN.ChartAreas(0).AxisX.LabelStyle.Angle = 90
        ChartBAYN.ChartAreas(0).AxisY.Interval = 50
        ChartBAYN.ChartAreas(0).AxisY.Minimum = 0
        ChartBAYN.Series(0).Points.AddXY(counter, BAYN)
        If Variation1 < 0 Then
            PriceBAYN.ForeColor = System.Drawing.Color.Red
            PriceBAYN.Text = "(" & Variation1 & "$) " & BAYN
        Else
            PriceBAYN.ForeColor = System.Drawing.Color.Green
            PriceBAYN.Text = "(+" & Variation1 & "$) " & BAYN
        End If
        'PriceBAYN.Text = "(" & Variation1 & "$) " & BAYN
        FB += Variation2
        ChartFB.ChartAreas(0).AxisX.Interval = 1
        ChartFB.ChartAreas(0).AxisX.Minimum = 0
        ChartFB.ChartAreas(0).AxisX.LabelStyle.Angle = 90
        ChartFB.ChartAreas(0).AxisY.Interval = 50
        ChartFB.ChartAreas(0).AxisY.Minimum = 0
        ChartFB.Series(0).Points.AddXY(counter, FB)
        If Variation2 < 0 Then
            PriceFB.ForeColor = System.Drawing.Color.Red
            PriceFB.Text = "(" & Variation2 & "$) " & FB
        Else
            PriceFB.ForeColor = System.Drawing.Color.Green
            PriceFB.Text = "(+" & Variation2 & "$) " & FB
        End If
        'PriceFB.Text = "(" & Variation2 & "$) " & FB
        GOOGL += Variation3
        ChartGOOGL.ChartAreas(0).AxisX.Interval = 1
        ChartGOOGL.ChartAreas(0).AxisX.Minimum = 0
        ChartGOOGL.ChartAreas(0).AxisX.LabelStyle.Angle = 90
        ChartGOOGL.ChartAreas(0).AxisY.Interval = 500
        ChartGOOGL.ChartAreas(0).AxisY.Minimum = 0
        ChartGOOGL.Series(0).Points.AddXY(counter, GOOGL)
        If Variation3 < 0 Then
            PriceGOOGL.ForeColor = System.Drawing.Color.Red
            PriceGOOGL.Text = "(" & Variation3 & "$) " & GOOGL
        Else
            PriceGOOGL.ForeColor = System.Drawing.Color.Green
            PriceGOOGL.Text = "(+" & Variation3 & "$) " & GOOGL
        End If
        'PriceGOOGL.Text = "(" & Variation3 & "$) " & GOOGL
        MSFT += Variation4
        ChartMSFT.ChartAreas(0).AxisX.Interval = 1
        ChartMSFT.ChartAreas(0).AxisX.Minimum = 0
        ChartMSFT.ChartAreas(0).AxisX.LabelStyle.Angle = 90
        ChartMSFT.ChartAreas(0).AxisY.Interval = 50
        ChartMSFT.ChartAreas(0).AxisY.Minimum = 0
        ChartMSFT.Series(0).Points.AddXY(counter, MSFT)
        If Variation4 < 0 Then
            PriceMSFT.ForeColor = System.Drawing.Color.Red
            PriceMSFT.Text = "(" & Variation4 & "$) " & MSFT
        Else
            PriceMSFT.ForeColor = System.Drawing.Color.Green
            PriceMSFT.Text = "(+" & Variation4 & "$) " & MSFT
        End If
        'PriceMSFT.Text = "(" & Variation4 & "$) " & MSFT
        AMZN += Variation5
        ChartAMZN.ChartAreas(0).AxisX.Interval = 1
        ChartAMZN.ChartAreas(0).AxisX.Minimum = 0
        ChartAMZN.ChartAreas(0).AxisX.LabelStyle.Angle = 90
        ChartAMZN.ChartAreas(0).AxisY.Interval = 1000
        ChartAMZN.ChartAreas(0).AxisY.Minimum = 0
        ChartAMZN.Series(0).Points.AddXY(counter, AMZN)
        If Variation5 < 0 Then
            PriceAMZN.ForeColor = System.Drawing.Color.Red
            PriceAMZN.Text = "(" & Variation5 & "$) " & AMZN
        Else
            PriceAMZN.ForeColor = System.Drawing.Color.Green
            PriceAMZN.Text = "(+" & Variation5 & "$) " & AMZN
        End If
        'PriceAMZN.Text = "(" & Variation5 & "$) " & AMZN
        NOVN += Variation6
        ChartNOVN.ChartAreas(0).AxisX.Interval = 1
        ChartNOVN.ChartAreas(0).AxisX.Minimum = 0
        ChartNOVN.ChartAreas(0).AxisX.LabelStyle.Angle = 90
        ChartNOVN.ChartAreas(0).AxisY.Interval = 50
        ChartNOVN.ChartAreas(0).AxisY.Minimum = 0
        ChartNOVN.Series(0).Points.AddXY(counter, NOVN)
        If Variation6 < 0 Then
            PriceNOVN.ForeColor = System.Drawing.Color.Red
            PriceNOVN.Text = "(" & Variation5 & "$) " & NOVN
        Else
            PriceNOVN.ForeColor = System.Drawing.Color.Green
            PriceNOVN.Text = "(+" & Variation5 & "$) " & NOVN
        End If
        'PriceNOVN.Text = "(" & Variation5 & "$) " & NOVN
        UBS += Variation7
        ChartUBS.ChartAreas(0).AxisX.Interval = 1
        ChartUBS.ChartAreas(0).AxisX.Minimum = 0
        ChartUBS.ChartAreas(0).AxisX.LabelStyle.Angle = 90
        ChartUBS.ChartAreas(0).AxisY.Interval = 50
        ChartUBS.ChartAreas(0).AxisY.Minimum = 0
        ChartUBS.Series(0).Points.AddXY(counter, UBS)
        If Variation7 < 0 Then
            PriceUBS.ForeColor = System.Drawing.Color.Red
            PriceUBS.Text = "(" & Variation6 & "$) " & UBS
        Else
            PriceUBS.ForeColor = System.Drawing.Color.Green
            PriceUBS.Text = "(+" & Variation6 & "$) " & UBS
        End If
        'PriceUBS.Text = "(" & Variation6 & "$) " & UBS
        counter += 1
    End Sub

    Private Sub Form2_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Form1.Hide()
    End Sub

    Private Sub PictureBox4_Click(sender As Object, e As EventArgs) Handles PictureBox4.Click

        If balance > -100 Then
            If TabControl1.SelectedTab.Text.Contains("AAPL") Then
                If balance >= AAPL Then
                    balance -= AAPL
                    Label3.Text = balance
                    If balance < 0 Then
                        Label3.ForeColor = System.Drawing.Color.Red
                    Else
                        Label3.ForeColor = System.Drawing.Color.Green
                    End If
                    ListBox1.Items.Add(TabControl1.SelectedTab.Text)
                    ListBox1.SelectedIndex = 0
                Else
                    MsgBox("Balance too low to make a transaction", vbCritical)
                End If

            ElseIf TabControl1.SelectedTab.Text.Contains("BAYN") Then
                If balance >= BAYN Then
                    balance -= BAYN
                    Label3.Text = balance
                    If balance < 0 Then
                        Label3.ForeColor = System.Drawing.Color.Red
                    Else
                        Label3.ForeColor = System.Drawing.Color.Green
                    End If
                    ListBox1.Items.Add(TabControl1.SelectedTab.Text)
                    ListBox1.SelectedIndex = 0
                Else
                    MsgBox("Balance too low to make a transaction", vbCritical)
                End If

            ElseIf TabControl1.SelectedTab.Text.Contains("MSFT") Then
                If balance >= MSFT Then
                    balance -= MSFT
                    Label3.Text = balance
                    If balance < 0 Then
                        Label3.ForeColor = System.Drawing.Color.Red
                    Else
                        Label3.ForeColor = System.Drawing.Color.Green
                    End If
                    ListBox1.Items.Add(TabControl1.SelectedTab.Text)
                    ListBox1.SelectedIndex = 0
                Else
                    MsgBox("Balance too low to make a transaction", vbCritical)
                End If

            ElseIf TabControl1.SelectedTab.Text.Contains("GOOGL") Then
                If balance >= GOOGL Then
                    balance -= GOOGL
                    Label3.Text = balance
                    If balance < 0 Then
                        Label3.ForeColor = System.Drawing.Color.Red
                    Else
                        Label3.ForeColor = System.Drawing.Color.Green
                    End If
                    ListBox1.Items.Add(TabControl1.SelectedTab.Text)
                    ListBox1.SelectedIndex = 0
                Else
                    MsgBox("Balance too low to make a transaction", vbCritical)
                End If

            ElseIf TabControl1.SelectedTab.Text.Contains("NOVN") Then
                If balance >= NOVN Then
                    balance -= NOVN
                    Label3.Text = balance
                    If balance < 0 Then
                        Label3.ForeColor = System.Drawing.Color.Red
                    Else
                        Label3.ForeColor = System.Drawing.Color.Green
                    End If
                    ListBox1.Items.Add(TabControl1.SelectedTab.Text)
                    ListBox1.SelectedIndex = 0
                Else
                    MsgBox("Balance too low to make a transaction", vbCritical)
                End If


            ElseIf TabControl1.SelectedTab.Text.Contains("AMZN") Then
                If balance >= AMZN Then
                    balance -= AMZN
                    Label3.Text = balance
                    If balance < 0 Then
                        Label3.ForeColor = System.Drawing.Color.Red
                    Else
                        Label3.ForeColor = System.Drawing.Color.Green
                    End If
                    ListBox1.Items.Add(TabControl1.SelectedTab.Text)
                    ListBox1.SelectedIndex = 0
                Else
                    MsgBox("Balance too low to make a transaction", vbCritical)
                End If

            ElseIf TabControl1.SelectedTab.Text.Contains("UBS") Then
                If balance >= UBS Then
                    balance -= UBS
                    Label3.Text = balance
                    If balance < 0 Then
                        Label3.ForeColor = System.Drawing.Color.Red
                    Else
                        Label3.ForeColor = System.Drawing.Color.Green
                    End If
                    ListBox1.Items.Add(TabControl1.SelectedTab.Text)
                    ListBox1.SelectedIndex = 0
                Else
                    MsgBox("Balance too low to make a transaction", vbCritical)
                End If

            ElseIf TabControl1.SelectedTab.Text.Contains("FB") Then
                If balance >= UBS Then
                    balance -= FB
                    Label3.Text = balance
                    If balance < 0 Then
                        Label3.ForeColor = System.Drawing.Color.Red
                    Else
                        Label3.ForeColor = System.Drawing.Color.Green
                    End If
                    ListBox1.Items.Add(TabControl1.SelectedTab.Text)
                    ListBox1.SelectedIndex = 0
                Else
                    MsgBox("Balance too low to make a transaction", vbCritical)
                End If

            End If

        Else
            MsgBox("Balance too low to make a transaction", vbCritical)
        End If


    End Sub

    Private Sub PictureBox5_Click(sender As Object, e As EventArgs) Handles PictureBox5.Click
        Try
            If ListBox1.SelectedItem.Contains("AAPL") Then
                balance += AAPL
                Label3.Text = balance
                If balance < 0 Then
                    Label3.ForeColor = System.Drawing.Color.Red
                Else
                    Label3.ForeColor = System.Drawing.Color.Green
                End If
                ListBox1.Items.RemoveAt(ListBox1.SelectedIndex)
            ElseIf ListBox1.SelectedItem.Contains("AMZN") Then
                balance += AMZN
                Label3.Text = balance
                If balance < 0 Then
                    Label3.ForeColor = System.Drawing.Color.Red
                Else
                    Label3.ForeColor = System.Drawing.Color.Green
                End If
                ListBox1.Items.RemoveAt(ListBox1.SelectedIndex)
            ElseIf ListBox1.SelectedItem.Contains("GOOGL") Then
                balance += GOOGL
                Label3.Text = balance
                If balance < 0 Then
                    Label3.ForeColor = System.Drawing.Color.Red
                Else
                    Label3.ForeColor = System.Drawing.Color.Green
                End If
                ListBox1.Items.RemoveAt(ListBox1.SelectedIndex)
            ElseIf ListBox1.SelectedItem.Contains("BAYN") Then
                balance += BAYN
                Label3.Text = balance
                If balance < 0 Then
                    Label3.ForeColor = System.Drawing.Color.Red
                Else
                    Label3.ForeColor = System.Drawing.Color.Green
                End If
                ListBox1.Items.RemoveAt(ListBox1.SelectedIndex)
            ElseIf ListBox1.SelectedItem.Contains("FB") Then
                balance += FB
                Label3.Text = balance
                If balance < 0 Then
                    Label3.ForeColor = System.Drawing.Color.Red
                Else
                    Label3.ForeColor = System.Drawing.Color.Green
                End If
                ListBox1.Items.RemoveAt(ListBox1.SelectedIndex)
            ElseIf ListBox1.SelectedItem.Contains("UBS") Then
                balance += UBS
                Label3.Text = balance
                If balance < 0 Then
                    Label3.ForeColor = System.Drawing.Color.Red
                Else
                    Label3.ForeColor = System.Drawing.Color.Green
                End If
                ListBox1.Items.RemoveAt(ListBox1.SelectedIndex)
            ElseIf ListBox1.SelectedItem.Contains("NOVN") Then
                balance += NOVN
                Label3.Text = balance
                If balance < 0 Then
                    Label3.ForeColor = System.Drawing.Color.Red
                Else
                    Label3.ForeColor = System.Drawing.Color.Green
                End If
                ListBox1.Items.RemoveAt(ListBox1.SelectedIndex)
            ElseIf ListBox1.SelectedItem.Contains("MSFT") Then
                balance += MSFT
                Label3.Text = balance
                If balance < 0 Then
                    Label3.ForeColor = System.Drawing.Color.Red
                Else
                    Label3.ForeColor = System.Drawing.Color.Green
                End If
                ListBox1.Items.RemoveAt(ListBox1.SelectedIndex)
            End If
        Catch ex6 As NullReferenceException
            MsgBox("No selected Stock", vbCritical)
        End Try

    End Sub

    Private Sub Timer2_Tick(sender As Object, e As EventArgs) Handles Timer2.Tick
        Label6.Text = TimeOfDay.ToString("h:mm:ss tt")
    End Sub
End Class