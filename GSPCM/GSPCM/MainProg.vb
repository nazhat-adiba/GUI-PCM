Imports System
Imports System.IO.Ports
Imports System.Text
Imports System.Math

Public Class MainProg
    Public ReceivedData As String
    Dim savee As String
    Dim save1 As String
    Dim save2 As String
    Dim iax As Double = 0
    Dim iay As Double = 0
    Dim iaz As Double = 0
    Dim data As String
    Dim err As Integer = 0
    Dim con As Boolean = False
    Dim strdata() As String
    Dim ax, ay, az, gxx, gyy, gzz, mgx, mgy, mgz, latitude, longitude, altitude As String
    Dim Datacount As String
    Dim vax0 As Double = 0
    Dim vay0 As Double = 0
    Dim vaz0 As Double = 0
    Dim bax As Boolean = False
    Dim bay As Boolean = False
    Dim baz As Boolean = False
    Dim bgx As Boolean = False
    Dim bgy As Boolean = False
    Dim bgz As Boolean = False
    Dim bmx As Boolean = False
    Dim bmy As Boolean = False
    Dim bmz As Boolean = False
    Dim blat As Boolean = False
    Dim blon As Boolean = False
    Dim balt As Boolean = False
    Dim state As Boolean = False
    Dim Sensorvalue As String
    Dim bin As String
    Dim SensVal As Integer
    Dim sensss As Double
    Dim SensStrVal As String
    Dim Header As String
    Dim strval As String
    Dim Vel As Double
    Dim vax As Double
    Dim vay As Double
    Dim vaz As Double
    Dim vaxs As String
    Dim vays As String
    Dim vazs As String

    Private Sub MainProg_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        For Each AvailablePorts As String In IO.Ports.SerialPort.GetPortNames()
            ComboBoxCom.Items.Add(AvailablePorts)
            SerialPort1.ReadTimeout = 50
        Next
        ComboBoxBaud.Items.Add(9600)
        ComboBoxBaud.Items.Add(57600)
        ComboBoxPar.Items.Add(Parity.None)
        ComboBoxPar.Items.Add(Parity.Even)
        ComboBoxPar.Items.Add(Parity.Odd)
        Timer1.Enabled = False
    End Sub

    Public Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If (Button1.Text = "Connect") Then
            con = True
            Button1.Text = "Disconnect"
        Else
            con = False
        End If
    End Sub

    Private Sub Data_Filter()
        RichTextBox1.Text &= SerialPort1.ReadLine
        data = SerialPort1.ReadLine
        Dim String2() As String
        String2 = Split(data, ";")
        If String2.Count = 14 Then
            If String2(0).Length = 12 Then
                ReceivedData = data
            End If
        End If
        'For U = 0 To RichTextBox1.Lines.Count - 2
        '    Dim String1 As String
        '    Dim String2() As String
        '    String1 = RichTextBox1.Lines(U)
        '    String2 = Split(String1, ";")
        '    If String2.Count = 14 Then
        '        If String2(0).Length = 12 Then
        '            ReceivedData = String1
        '        End If
        '    End If
        'Next
    End Sub

    Private Sub Data_Splitter()
        If Not ReceivedData = Nothing Then
            strdata = ReceivedData.Split(";")
            Header = strdata(0)
            gxx = strdata(1)
            gyy = strdata(2)
            gzz = strdata(3)
            ax = strdata(4)
            ay = strdata(5)
            az = strdata(6)
            mgx = strdata(7)
            mgy = strdata(8)
            mgz = strdata(9)
            latitude = strdata(10)
            longitude = strdata(11)
            altitude = strdata(12)
            Datacount = strdata(13)
            If ReceivedData = Nothing Then
            End If
        End If
    End Sub

    Private Sub Header_Scanner()
        If Header.Substring(0, 1) = "1" Then
            bgx = True
        Else
            bgx = False
        End If
        If Header.Substring(1, 1) = "1" Then
            bgy = True
        Else
            bgy = False
        End If
        If Header.Substring(2, 1) = "1" Then
            bgz = True
        Else
            bgz = False
        End If
        If Header.Substring(3, 1) = "1" Then
            bax = True
        Else
            bax = False
        End If
        If Header.Substring(4, 1) = "1" Then
            bay = True
        Else
            bay = False
        End If
        If Header.Substring(5, 1) = "1" Then
            baz = True
        Else
            baz = False
        End If
        If Header.Substring(6, 1) = "1" Then
            bmx = True
        Else
            bmx = False
        End If
        If Header.Substring(7, 1) = "1" Then
            bmy = True
        Else
            bmy = False
        End If
        If Header.Substring(8, 1) = "1" Then
            bmz = True
        Else
            bmz = False
        End If
        If Header.Substring(9, 1) = "1" Then
            blat = True
        Else
            blat = False
        End If
        If Header.Substring(10, 1) = "1" Then
            blon = True
        Else
            blon = False
        End If
        If Header.Substring(11, 1) = "1" Then
            balt = True
        Else
            balt = False
        End If
    End Sub

    Private Sub PCM_Binary_Scanner()
        Dim bit0 As Integer = 0
        Dim bit1 As Integer = 0
        Dim bit2 As Integer = 0
        Dim bit3 As Integer = 0
        Dim bit4 As Integer = 0
        Dim bit5 As Integer = 0
        Dim bit6 As Integer = 0
        Dim bit7 As Integer = 0
        Dim bit8 As Integer = 0
        Dim bit9 As Integer = 0
        Dim bit10 As Integer = 0
        Dim bit11 As Integer = 0
        Dim bit12 As Integer = 0
        Dim bit13 As Integer = 0
        Dim bit14 As Integer = 0
        Dim bit15 As Integer = 0
        Dim ch0, ch1, ch2, ch3, ch4, ch5, ch6, ch7, ch8, ch9, ch10, ch11, ch12, ch13, ch14, ch15 As String
        'counter program
        SensVal = 0
        If Sensorvalue.IndexOf("0000000000000000") = -1 Then
            For DD = 1 To 65536
                bit0 += 1
                If bit0 = 2 Then
                    bit1 += 1
                    bit0 = 0
                    If bit1 = 2 Then
                        bit2 += 1
                        bit1 = 0
                        If bit2 = 2 Then
                            bit3 += 1
                            bit2 = 0
                            If bit3 = 2 Then
                                bit4 += 1
                                bit3 = 0
                                If bit4 = 2 Then
                                    bit5 += 1
                                    bit4 = 0
                                    If bit5 = 2 Then
                                        bit6 += 1
                                        bit5 = 0
                                        If bit6 = 2 Then
                                            bit7 += 1
                                            bit6 = 0
                                            If bit7 = 2 Then
                                                bit8 += 1
                                                bit7 = 0
                                                If bit8 = 2 Then
                                                    bit9 += 1
                                                    bit8 = 0
                                                    If bit9 = 2 Then
                                                        bit10 += 1
                                                        bit9 = 0
                                                        If bit10 = 2 Then
                                                            bit11 += 1
                                                            bit10 = 0
                                                            If bit11 = 2 Then
                                                                bit12 += 1
                                                                bit11 = 0
                                                                If bit12 = 2 Then
                                                                    bit13 += 1
                                                                    bit12 = 0
                                                                    If bit13 = 2 Then
                                                                        bit14 += 1
                                                                        bit13 = 0
                                                                        If bit14 = 2 Then
                                                                            bit15 += 1
                                                                            bit14 = 0
                                                                            If bit15 = 2 Then
                                                                                bit15 = 1
                                                                            End If
                                                                        End If
                                                                    End If
                                                                End If
                                                            End If
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
                ch0 = bit0
                ch1 = bit1
                ch2 = bit2
                ch3 = bit3
                ch4 = bit4
                ch5 = bit5
                ch6 = bit6
                ch7 = bit7
                ch8 = bit8
                ch9 = bit9
                ch10 = bit10
                ch11 = bit11
                ch12 = bit12
                ch13 = bit13
                ch14 = bit14
                ch15 = bit15
                bin = ch15 + ch14 + ch13 + ch12 + ch11 + ch10 + ch9 + ch8 + ch7 + ch6 + ch5 + ch4 + ch3 + ch2 + ch1 + ch0
                If Not Sensorvalue.IndexOf(bin) = -1 Then
                    SensVal = DD
                    Exit For
                End If
            Next
        Else
            SensVal = 0
        End If
    End Sub
    'Private Sub PCM_Decoder()
    '    SensVal = 0
    '    If Not Sensorvalue.IndexOf("0000000000000000") = -1 Then
    '        For DD = 1 To 65536
    '            PCM_Binary_Scanner()
    '            If Not Sensorvalue.IndexOf(bin) = -1 Then
    '                SensVal = DD
    '                Exit For
    '            End If
    '        Next
    '    Else
    '        SensVal = 0
    '    End If
    'End Sub
    Private Sub Conversion()
        sensss = SensVal / 10
        SensStrVal = sensss
    End Sub

    Private Sub Value_Scanner()
        Header_Scanner()
        If state = True Then
            PCM_Binary_Scanner()
            Conversion()
            strval = SensStrVal
        Else
            PCM_Binary_Scanner()
            Conversion()
            strval = "-" + SensStrVal
        End If
    End Sub

    'Private Sub Velocity_Measuerement()
    '    iax = 0
    '    iay = 0
    '    iaz = 0
    '    iax = Val(vaxs) / 0.101
    '    iay = Val(vays) / 0.101
    '    iaz = Val(vazs) / 0.101
    '    vax = (iax * 0.1) + vax0
    '    vax0 = vax
    '    vay = (iay * 0.1) + vay0
    '    vay0 = vay
    '    vaz = (iaz * 0.1) + vaz0
    '    vaz0 = vaz
    '    Vel = Sqrt((vax * vax) + (vay * vay) + (vaz * vaz))
    '    Label22.Text = Vel
    'End Sub

    Private Sub graphic()
        Chart1.Series("Series1").Points.Add(ACCX.Text)
        If Chart1.Series(0).Points.Count = 15 Then
            Chart1.Series(0).Points.RemoveAt(0)
        End If
        Chart1.Series("Series2").Points.Add(ACCY.Text)
        If Chart1.Series(1).Points.Count = 15 Then
            Chart1.Series(1).Points.RemoveAt(0)
        End If
        Chart1.Series("Series3").Points.Add(ACCZ.Text)
        If Chart1.Series(2).Points.Count = 15 Then
            Chart1.Series(2).Points.RemoveAt(0)
        End If
        Chart2.Series("Series1").Points.Add(GX.Text)
        If Chart2.Series(0).Points.Count = 15 Then
            Chart2.Series(0).Points.RemoveAt(0)
        End If
        Chart2.Series("Series2").Points.Add(GY.Text)
        If Chart2.Series(1).Points.Count = 15 Then
            Chart2.Series(1).Points.RemoveAt(0)
        End If
        Chart2.Series("Series3").Points.Add(GZ.Text)
        If Chart2.Series(2).Points.Count = 15 Then
            Chart2.Series(2).Points.RemoveAt(0)
        End If
        Chart3.Series("Series1").Points.Add(MX.Text)
        If Chart3.Series(0).Points.Count = 15 Then
            Chart3.Series(0).Points.RemoveAt(0)
        End If
        Chart3.Series("Series2").Points.Add(MY2.Text)
        If Chart3.Series(1).Points.Count = 15 Then
            Chart3.Series(1).Points.RemoveAt(0)
        End If
        Chart3.Series("Series3").Points.Add(MZ.Text)
        If Chart3.Series(2).Points.Count = 15 Then
            Chart3.Series(2).Points.RemoveAt(0)
        End If
        Chart4.Series("Series1").Points.Add(LAT.Text)
        If Chart4.Series(0).Points.Count = 15 Then
            Chart4.Series(0).Points.RemoveAt(0)
        End If
        Chart4.Series("Series2").Points.Add(LONGT.Text)
        If Chart4.Series(1).Points.Count = 15 Then
            Chart4.Series(1).Points.RemoveAt(0)
        End If
        Chart4.Series("Series3").Points.Add(ALT.Text)
        If Chart4.Series(2).Points.Count = 15 Then
            Chart4.Series(2).Points.RemoveAt(0)
        End If
    End Sub

    Private Sub rotation()

    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        Data_Filter()
        If Not ReceivedData = Nothing Then
            Data_Splitter()
            'Accelerometer Process
            state = bax
            Sensorvalue = ax
            Value_Scanner()
            vaxs = strval
            ACCX.Text = strval
            state = bay
            Sensorvalue = ay
            Value_Scanner()
            vays = strval
            ACCY.Text = strval
            state = baz
            Sensorvalue = az
            Value_Scanner()
            vazs = strval
            ACCZ.Text = strval
            'Gyroscope Process
            state = bgx
            Sensorvalue = gxx
            Value_Scanner()
            GX.Text = strval
            state = bgy
            Sensorvalue = gyy
            Value_Scanner()
            GY.Text = strval
            state = bgz
            Sensorvalue = gzz
            Value_Scanner()
            GZ.Text = strval
            'Magnetometer Process
            state = bmx
            Sensorvalue = mgx
            Value_Scanner()
            MX.Text = strval
            state = bmy
            Sensorvalue = mgy
            Value_Scanner()
            MY2.Text = strval
            state = bmz
            Sensorvalue = mgz
            Value_Scanner()
            MZ.Text = strval
            'GPS Process
            state = blat
            Sensorvalue = latitude
            Value_Scanner()
            LAT.Text = strval
            state = blon
            Sensorvalue = longitude
            Value_Scanner()
            LONGT.Text = strval
            state = balt
            Sensorvalue = altitude
            Value_Scanner()
            ALT.Text = strval
            graphic()
            'Velocity_Measuerement()
            savee = ACCX.Text + ";" + ACCY.Text + ";" + ACCZ.Text + ";" + GX.Text + ";" + GY.Text + ";" + GZ.Text + ";" + MX.Text + ";" + MY2.Text + ";" + MZ.Text + ";" + LAT.Text + ";" + LONGT.Text + ";" + ALT.Text + ";" + Datacount
            'RichTextBox2.Text = savee
            save1 = ReceivedData
            If CheckBox1.Checked Then
                My.Computer.FileSystem.WriteAllText(TextBox1.Text, savee + vbCrLf, True)
                My.Computer.FileSystem.WriteAllText(TextBox2.Text, save1 + vbCrLf, True)
            End If
            If ReceivedData = Nothing Then
            End If
        End If

    End Sub

    Private Sub RichTextBox1_TextChanged(sender As Object, e As EventArgs) Handles RichTextBox1.TextChanged
        If RichTextBox1.Lines.Count = 10 Then
            RichTextBox1.Select(0, RichTextBox1.GetFirstCharIndexFromLine(1))
            RichTextBox1.SelectedText = ""
        End If
        RichTextBox1.SelectionStart = RichTextBox1.TextLength
        RichTextBox1.ScrollToCaret()
    End Sub

    Private Sub SaveFileDialog1_FileOk(sender As Object, e As ComponentModel.CancelEventArgs) Handles SaveFileDialog1.FileOk

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If con = True Then
            If (Button2.Text = "Start") Then
                SerialPort1.Close()
                SerialPort1.PortName = ComboBoxCom.SelectedItem
                SerialPort1.BaudRate = ComboBoxBaud.SelectedItem
                SerialPort1.Parity = ComboBoxPar.SelectedItem
                SerialPort1.ReadTimeout = 500
                SerialPort1.Open()
                Button2.Text = "Stop"
                If SerialPort1.IsOpen Then
                    Timer1.Enabled = True
                Else
                    Timer1.Enabled = False
                    MsgBox("Device Disconnected")
                    Button1.Text = "Start"
                End If
            Else
                Timer1.Enabled = False
                SerialPort1.Close()
                Button1.Text = "Start"
            End If
        Else
            Timer1.Enabled = False
            SerialPort1.Close()
        End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        SaveFileDialog1.Filter = "TXT FILE|*.txt"
        SaveFileDialog1.FileName = "DataSensorLog"
        If SaveFileDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
            TextBox1.Text = SaveFileDialog1.FileName
        End If
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        SaveFileDialog2.Filter = "TXT FILE|*.txt"
        SaveFileDialog2.FileName = "DataBinary"
        If SaveFileDialog2.ShowDialog = Windows.Forms.DialogResult.OK Then
            TextBox2.Text = SaveFileDialog2.FileName
        End If
    End Sub
End Class