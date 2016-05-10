Public Class SC16A
    Dim WithEvents serialPort As New IO.Ports.SerialPort

    Private Const AW_BLEND = &H80000  'Uses a fade effect. This flag can be used only if hwnd is a top-level window.
    Private Declare Function AnimateWindow Lib "user32" (ByVal hwnd As Int32, ByVal dwTime As Int32, ByVal dwFlags As Int32) As Boolean
    Dim winHide As Integer = &H10000
    Dim winBlend As Integer = &H80000

    Dim servo1 As Integer
    Dim servo2 As Integer
    Dim servo3 As Integer
    Dim servo4 As Integer
    Dim servo5 As Integer
    Dim servo6 As Integer
    Dim servo7 As Integer
    Dim servo8 As Integer
    Dim servo9 As Integer
    Dim servo10 As Integer
    Dim servo11 As Integer
    Dim servo12 As Integer
    Dim servo13 As Integer
    Dim servo14 As Integer
    Dim servo15 As Integer
    Dim servo16 As Integer
    Dim servo17 As Integer
    Dim servo18 As Integer
    Dim servo19 As Integer
    Dim servo20 As Integer
    Dim servo21 As Integer
    Dim servo22 As Integer
    Dim servo23 As Integer
    Dim servo24 As Integer
    Dim servo25 As Integer
    Dim servo26 As Integer
    Dim servo27 As Integer
    Dim servo28 As Integer
    Dim servo29 As Integer
    Dim servo30 As Integer
    Dim servo31 As Integer
    Dim servo32 As Integer
    Dim data1 As Byte       'higher byte
    Dim data2 As Byte       'lower byte
    Dim data3 As Byte       'ramp/speed value
    Dim pattern As Integer
    Dim pattern2 As Integer
    Dim pattern3 As Integer
    Dim pattern4 As Integer
    Dim pattern5 As Integer
    Dim pattern6 As Integer
    Dim pattern7 As Integer
    Dim pattern8 As Integer
    Dim pattern9 As Integer
    Dim pattern10 As Integer
    Dim pattern11 As Integer
    Dim pattern12 As Integer
    Dim pattern13 As Integer
    Dim pattern14 As Integer
    Dim pattern15 As Integer
    Dim pattern16 As Integer
    Dim pattern17 As Integer
    Dim pattern18 As Integer
    Dim pattern19 As Integer
    Dim pattern20 As Integer
    Dim pattern21 As Integer
    Dim pattern22 As Integer
    Dim pattern23 As Integer
    Dim pattern24 As Integer
    Dim pattern25 As Integer
    Dim pattern26 As Integer
    Dim pattern27 As Integer
    Dim pattern28 As Integer
    Dim pattern29 As Integer
    Dim pattern30 As Integer
    Dim pattern31 As Integer
    Dim pattern32 As Integer
    Dim s_en As Integer

    Private Sub SC16A_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed

        Try
            serialPort.Close()
            s_en = 0
            AnimateWindow(Me.Handle.ToInt32, CInt(800), winHide Or winBlend)
        Catch ex As Exception
           
        End Try

    End Sub


    Private Sub SC16A_Load( _
       ByVal sender As System.Object, _
       ByVal e As System.EventArgs) _
       Handles MyBase.Load

        For i As Integer = 0 To _
           My.Computer.Ports.SerialPortNames.Count - 1
            cbbCOMPorts.Items.Add( _
               My.Computer.Ports.SerialPortNames(i))
        Next
        btnDisconnect.Enabled = False
        pattern = 0
        pattern2 = 0
        pattern3 = 0
        pattern4 = 0
        pattern5 = 0
        pattern6 = 0
        pattern7 = 0
        pattern8 = 0
        pattern9 = 0
        pattern10 = 0
        pattern11 = 0
        pattern12 = 0
        pattern13 = 0
        pattern14 = 0
        pattern15 = 0
        pattern16 = 0
        s_en = 0

    End Sub


    Private Sub DataReceived( _
       ByVal sender As Object, _
       ByVal e As System.IO.Ports.SerialDataReceivedEventArgs) _
       Handles serialPort.DataReceived

        txtDataReceived.Invoke(New _
                       myDelegate(AddressOf updateTextBox), _
                       New Object() {})
    End Sub

    Public Delegate Sub myDelegate()
    Public Sub updateTextBox()
        With txtDataReceived
            .Font = New Font("Garamond", 12.0!, FontStyle.Bold)
            .SelectionColor = Color.Red
            .AppendText(serialPort.ReadExisting)
            .ScrollToCaret()
        End With
    End Sub

    Private Sub btnConnect_Click( _
       ByVal sender As System.Object, _
       ByVal e As System.EventArgs) _
       Handles btnConnect.Click
        If serialPort.IsOpen Then
            serialPort.Close()
        End If
        Try
            With serialPort
                .PortName = cbbCOMPorts.Text
                .BaudRate = 9600 '115200
                .Parity = IO.Ports.Parity.None
                .DataBits = 8
                .StopBits = IO.Ports.StopBits.One
                ' .Encoding = System.Text.Encoding.Unicode
            End With
            serialPort.Open()

            lblMessage.Text = cbbCOMPorts.Text & " connected."
            btnConnect.Enabled = False
            btnDisconnect.Enabled = True
            s_en = 1
        Catch ex As Exception
            'MsgBox(ex.ToString)
            MsgBox("Please select COM PORTS")

        End Try
    End Sub

    Private Sub btnDisconnect_Click( _
       ByVal sender As System.Object, _
       ByVal e As System.EventArgs) _
       Handles btnDisconnect.Click
        Try
            serialPort.Close()
            lblMessage.Text = serialPort.PortName & " disconnected."
            btnConnect.Enabled = True
            btnDisconnect.Enabled = False
            s_en = 0
        Catch ex As Exception
            'MsgBox(ex.ToString)
            'MsgBox("unplugged")
            lblMessage.Text = serialPort.PortName & " disconnected."
            btnConnect.Enabled = True
            btnDisconnect.Enabled = False

        End Try

    End Sub
    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        cbbCOMPorts.Items.Clear()


        For i As Integer = 0 To _
          My.Computer.Ports.SerialPortNames.Count - 1
            cbbCOMPorts.Items.Add( _
               My.Computer.Ports.SerialPortNames(i))
        Next
        btnDisconnect.Enabled = False
    End Sub
   
    'servo 1
    Private Sub VScrollBar1_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles VScrollBar1.ValueChanged
        TextBox1.Text = VScrollBar1.Value
        servo1 = VScrollBar1.Value

        data1 = (servo1 >> 6) And &HFF
        data2 = servo1 And &H3F
        data3 = NumericUpDown1.Value

        If s_en = 1 Then
            Try
                serialPort.Write(StrConv(Chr(&H41), VbStrConv.None))
                serialPort.Write(StrConv(Chr(data1), VbStrConv.None))
                serialPort.Write(StrConv(Chr(data2), VbStrConv.None))
                serialPort.Write(StrConv(Chr(data3), VbStrConv.None))
              
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try

        End If
    End Sub
    'servo 2
    Private Sub VScrollBar2_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles VScrollBar2.ValueChanged
        TextBox2.Text = VScrollBar2.Value
        servo2 = VScrollBar2.Value

        data1 = (servo2 >> 6) And &HFF
        data2 = servo2 And &H3F
        data3 = NumericUpDown2.Value

        If s_en = 1 Then
            Try
                serialPort.Write(StrConv(Chr(&H42), VbStrConv.None))
                serialPort.Write(StrConv(Chr(data1), VbStrConv.None))
                serialPort.Write(StrConv(Chr(data2), VbStrConv.None))
                serialPort.Write(StrConv(Chr(data3), VbStrConv.None))
            
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try

        End If


    End Sub
    'servo 3
    Private Sub VScrollBar3_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles VScrollBar3.ValueChanged
        TextBox3.Text = VScrollBar3.Value
        servo3 = VScrollBar3.Value

        data1 = (servo3 >> 6) And &HFF
        data2 = servo3 And &H3F
        data3 = NumericUpDown3.Value

        If s_en = 1 Then
            Try
                serialPort.Write(StrConv(Chr(&H43), VbStrConv.None))
                serialPort.Write(StrConv(Chr(data1), VbStrConv.None))
                serialPort.Write(StrConv(Chr(data2), VbStrConv.None))
                serialPort.Write(StrConv(Chr(data3), VbStrConv.None))
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try

        End If


    End Sub
    'servo 4
    Private Sub VScrollBar4_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles VScrollBar4.ValueChanged
        TextBox4.Text = VScrollBar4.Value
        servo4 = VScrollBar4.Value

        data1 = (servo4 >> 6) And &HFF
        data2 = servo4 And &H3F
        data3 = NumericUpDown4.Value

        If s_en = 1 Then
            Try
                serialPort.Write(StrConv(Chr(&H44), VbStrConv.None))
                serialPort.Write(StrConv(Chr(data1), VbStrConv.None))
                serialPort.Write(StrConv(Chr(data2), VbStrConv.None))
                serialPort.Write(StrConv(Chr(data3), VbStrConv.None))
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try

        End If


    End Sub
    'servo 5
    Private Sub VScrollBar5_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles VScrollBar5.ValueChanged
        TextBox5.Text = VScrollBar5.Value
        servo5 = VScrollBar5.Value

        data1 = (servo5 >> 6) And &HFF
        data2 = servo5 And &H3F
        data3 = NumericUpDown5.Value

        If s_en = 1 Then
            Try
                serialPort.Write(StrConv(Chr(&H45), VbStrConv.None))
                serialPort.Write(StrConv(Chr(data1), VbStrConv.None))
                serialPort.Write(StrConv(Chr(data2), VbStrConv.None))
                serialPort.Write(StrConv(Chr(data3), VbStrConv.None))
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try

        End If


    End Sub
    'servo 6
    Private Sub VScrollBar6_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles VScrollBar6.ValueChanged
        TextBox6.Text = VScrollBar6.Value
        servo6 = VScrollBar6.Value

        data1 = (servo6 >> 6) And &HFF
        data2 = servo6 And &H3F
        data3 = NumericUpDown6.Value

        If s_en = 1 Then
            Try
                serialPort.Write(StrConv(Chr(&H46), VbStrConv.None))
                serialPort.Write(StrConv(Chr(data1), VbStrConv.None))
                serialPort.Write(StrConv(Chr(data2), VbStrConv.None))
                serialPort.Write(StrConv(Chr(data3), VbStrConv.None))
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try

        End If


    End Sub
    'servo 7
    Private Sub VScrollBar7_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles VScrollBar7.ValueChanged
        TextBox7.Text = VScrollBar7.Value
        servo7 = VScrollBar7.Value

        data1 = (servo7 >> 6) And &HFF
        data2 = servo7 And &H3F
        data3 = NumericUpDown7.Value

        If s_en = 1 Then
            Try
                serialPort.Write(StrConv(Chr(&H47), VbStrConv.None))
                serialPort.Write(StrConv(Chr(data1), VbStrConv.None))
                serialPort.Write(StrConv(Chr(data2), VbStrConv.None))
                serialPort.Write(StrConv(Chr(data3), VbStrConv.None))
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try

        End If


    End Sub
    'servo 8
    Private Sub VScrollBar8_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles VScrollBar8.ValueChanged
        TextBox8.Text = VScrollBar8.Value
        servo8 = VScrollBar8.Value

        data1 = (servo8 >> 6) And &HFF
        data2 = servo8 And &H3F
        data3 = NumericUpDown8.Value

        If s_en = 1 Then
            Try
                serialPort.Write(StrConv(Chr(&H48), VbStrConv.None))
                serialPort.Write(StrConv(Chr(data1), VbStrConv.None))
                serialPort.Write(StrConv(Chr(data2), VbStrConv.None))
                serialPort.Write(StrConv(Chr(data3), VbStrConv.None))
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try

        End If


    End Sub
    'servo 9
    Private Sub VScrollBar9_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles VScrollBar9.ValueChanged
        TextBox9.Text = VScrollBar9.Value
        servo9 = VScrollBar9.Value

        data1 = (servo9 >> 6) And &HFF
        data2 = servo9 And &H3F
        data3 = NumericUpDown9.Value

        If s_en = 1 Then
            Try
                serialPort.Write(StrConv(Chr(&H49), VbStrConv.None))
                serialPort.Write(StrConv(Chr(data1), VbStrConv.None))
                serialPort.Write(StrConv(Chr(data2), VbStrConv.None))
                serialPort.Write(StrConv(Chr(data3), VbStrConv.None))
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try

        End If


    End Sub
    'servo 10
    Private Sub VScrollBar10_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles VScrollBar10.ValueChanged
        TextBox10.Text = VScrollBar10.Value
        servo10 = VScrollBar10.Value

        data1 = (servo10 >> 6) And &HFF
        data2 = servo10 And &H3F
        data3 = NumericUpDown10.Value

        If s_en = 1 Then
            Try
                serialPort.Write(StrConv(Chr(&H4A), VbStrConv.None))
                serialPort.Write(StrConv(Chr(data1), VbStrConv.None))
                serialPort.Write(StrConv(Chr(data2), VbStrConv.None))
                serialPort.Write(StrConv(Chr(data3), VbStrConv.None))
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try

        End If


    End Sub
    'servo 11
    Private Sub VScrollBar11_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles VScrollBar11.ValueChanged
        TextBox11.Text = VScrollBar11.Value
        servo11 = VScrollBar11.Value

        data1 = (servo11 >> 6) And &HFF
        data2 = servo11 And &H3F
        data3 = NumericUpDown11.Value

        If s_en = 1 Then
            Try
                serialPort.Write(StrConv(Chr(&H4B), VbStrConv.None))
                serialPort.Write(StrConv(Chr(data1), VbStrConv.None))
                serialPort.Write(StrConv(Chr(data2), VbStrConv.None))
                serialPort.Write(StrConv(Chr(data3), VbStrConv.None))
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try

        End If


    End Sub
    'servo 12
    Private Sub VScrollBar12_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles VScrollBar12.ValueChanged
        TextBox12.Text = VScrollBar12.Value
        servo12 = VScrollBar12.Value

        data1 = (servo12 >> 6) And &HFF
        data2 = servo12 And &H3F
        data3 = NumericUpDown12.Value

        If s_en = 1 Then
            Try
                serialPort.Write(StrConv(Chr(&H4C), VbStrConv.None))
                serialPort.Write(StrConv(Chr(data1), VbStrConv.None))
                serialPort.Write(StrConv(Chr(data2), VbStrConv.None))
                serialPort.Write(StrConv(Chr(data3), VbStrConv.None))
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try

        End If


    End Sub
    'servo 13
    Private Sub VScrollBar13_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles VScrollBar13.ValueChanged
        TextBox13.Text = VScrollBar13.Value
        servo13 = VScrollBar13.Value

        data1 = (servo13 >> 6) And &HFF
        data2 = servo13 And &H3F
        data3 = NumericUpDown13.Value

        If s_en = 1 Then
            Try
                serialPort.Write(StrConv(Chr(&H4D), VbStrConv.None))
                serialPort.Write(StrConv(Chr(data1), VbStrConv.None))
                serialPort.Write(StrConv(Chr(data2), VbStrConv.None))
                serialPort.Write(StrConv(Chr(data3), VbStrConv.None))
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try

        End If


    End Sub
    'servo 14
    Private Sub VScrollBar14_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles VScrollBar14.ValueChanged
        TextBox14.Text = VScrollBar14.Value
        servo14 = VScrollBar14.Value

        data1 = (servo14 >> 6) And &HFF
        data2 = servo14 And &H3F
        data3 = NumericUpDown14.Value

        If s_en = 1 Then
            Try
                serialPort.Write(StrConv(Chr(&H4E), VbStrConv.None))
                serialPort.Write(StrConv(Chr(data1), VbStrConv.None))
                serialPort.Write(StrConv(Chr(data2), VbStrConv.None))
                serialPort.Write(StrConv(Chr(data3), VbStrConv.None))
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try

        End If


    End Sub
    'servo 15
    Private Sub VScrollBar15_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles VScrollBar15.ValueChanged
        TextBox15.Text = VScrollBar15.Value
        servo15 = VScrollBar15.Value

        data1 = (servo15 >> 6) And &HFF
        data2 = servo15 And &H3F
        data3 = NumericUpDown15.Value

        If s_en = 1 Then
            Try
                serialPort.Write(StrConv(Chr(&H4F), VbStrConv.None))
                serialPort.Write(StrConv(Chr(data1), VbStrConv.None))
                serialPort.Write(StrConv(Chr(data2), VbStrConv.None))
                serialPort.Write(StrConv(Chr(data3), VbStrConv.None))
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try

        End If


    End Sub
    'servo 16
    Private Sub VScrollBar16_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles VScrollBar16.ValueChanged
        TextBox16.Text = VScrollBar16.Value
        servo16 = VScrollBar16.Value

        data1 = (servo16 >> 6) And &HFF
        data2 = servo16 And &H3F
        data3 = NumericUpDown16.Value

        If s_en = 1 Then
            Try
                serialPort.Write(StrConv(Chr(&H50), VbStrConv.None))
                serialPort.Write(StrConv(Chr(data1), VbStrConv.None))
                serialPort.Write(StrConv(Chr(data2), VbStrConv.None))
                serialPort.Write(StrConv(Chr(data3), VbStrConv.None))

            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try

        End If


    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        VScrollBar1.Value = 731
        VScrollBar2.Value = 731
        VScrollBar3.Value = 731
        VScrollBar4.Value = 731
        VScrollBar5.Value = 731
        VScrollBar6.Value = 731
        VScrollBar7.Value = 731
        VScrollBar8.Value = 731
        VScrollBar9.Value = 731
        VScrollBar10.Value = 731
        VScrollBar11.Value = 731
        VScrollBar12.Value = 731
        VScrollBar13.Value = 731
        VScrollBar14.Value = 731
        VScrollBar15.Value = 731
        VScrollBar16.Value = 731
        NumericUpDown1.Value = 0
        NumericUpDown2.Value = 0
        NumericUpDown3.Value = 0
        NumericUpDown4.Value = 0
        NumericUpDown5.Value = 0
        NumericUpDown6.Value = 0
        NumericUpDown7.Value = 0
        NumericUpDown8.Value = 0
        NumericUpDown9.Value = 0
        NumericUpDown10.Value = 0
        NumericUpDown11.Value = 0
        NumericUpDown12.Value = 0
        NumericUpDown13.Value = 0
        NumericUpDown14.Value = 0
        NumericUpDown15.Value = 0
        NumericUpDown16.Value = 0

    End Sub
    '32 ch mode

    'Servo 17
    Private Sub VScrollBar17_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles VScrollBar17.ValueChanged
        TextBox32.Text = VScrollBar17.Value
        servo17 = VScrollBar17.Value

        data1 = (servo17 >> 6) And &HFF
        data2 = servo17 And &H3F
        data3 = NumericUpDown32.Value

        If s_en = 1 Then
            Try
                serialPort.Write(StrConv(Chr(&H51), VbStrConv.None))
                serialPort.Write(StrConv(Chr(data1), VbStrConv.None))
                serialPort.Write(StrConv(Chr(data2), VbStrConv.None))
                serialPort.Write(StrConv(Chr(data3), VbStrConv.None))
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try

        End If

    End Sub
    'Servo 18
    Private Sub VScrollBar20_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles VScrollBar20.ValueChanged
        TextBox31.Text = VScrollBar20.Value
        servo18 = VScrollBar20.Value

        data1 = (servo18 >> 6) And &HFF
        data2 = servo18 And &H3F
        data3 = NumericUpDown31.Value

        If s_en = 1 Then
            Try
                serialPort.Write(StrConv(Chr(&H52), VbStrConv.None))
                serialPort.Write(StrConv(Chr(data1), VbStrConv.None))
                serialPort.Write(StrConv(Chr(data2), VbStrConv.None))
                serialPort.Write(StrConv(Chr(data3), VbStrConv.None))
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try

        End If
    End Sub
    'Servo 19
    Private Sub VScrollBar22_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles VScrollBar22.ValueChanged
        TextBox30.Text = VScrollBar22.Value
        servo19 = VScrollBar22.Value

        data1 = (servo19 >> 6) And &HFF
        data2 = servo19 And &H3F
        data3 = NumericUpDown30.Value

        If s_en = 1 Then
            Try
                serialPort.Write(StrConv(Chr(&H53), VbStrConv.None))
                serialPort.Write(StrConv(Chr(data1), VbStrConv.None))
                serialPort.Write(StrConv(Chr(data2), VbStrConv.None))
                serialPort.Write(StrConv(Chr(data3), VbStrConv.None))
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try

        End If
    End Sub
    'servo 20
    Private Sub VScrollBar24_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles VScrollBar24.ValueChanged
        TextBox29.Text = VScrollBar24.Value
        servo20 = VScrollBar24.Value

        data1 = (servo20 >> 6) And &HFF
        data2 = servo20 And &H3F
        data3 = NumericUpDown29.Value

        If s_en = 1 Then
            Try
                serialPort.Write(StrConv(Chr(&H54), VbStrConv.None))
                serialPort.Write(StrConv(Chr(data1), VbStrConv.None))
                serialPort.Write(StrConv(Chr(data2), VbStrConv.None))
                serialPort.Write(StrConv(Chr(data3), VbStrConv.None))
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try

        End If
    End Sub

    'servo 21
    Private Sub VScrollBar26_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles VScrollBar26.ValueChanged
        TextBox28.Text = VScrollBar26.Value
        servo21 = VScrollBar26.Value

        data1 = (servo21 >> 6) And &HFF
        data2 = servo21 And &H3F
        data3 = NumericUpDown28.Value

        If s_en = 1 Then
            Try
                serialPort.Write(StrConv(Chr(&H55), VbStrConv.None))
                serialPort.Write(StrConv(Chr(data1), VbStrConv.None))
                serialPort.Write(StrConv(Chr(data2), VbStrConv.None))
                serialPort.Write(StrConv(Chr(data3), VbStrConv.None))
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try

        End If
    End Sub
    'servo 22

    Private Sub VScrollBar28_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles VScrollBar28.ValueChanged
        TextBox27.Text = VScrollBar28.Value
        servo22 = VScrollBar28.Value

        data1 = (servo22 >> 6) And &HFF
        data2 = servo22 And &H3F
        data3 = NumericUpDown27.Value

        If s_en = 1 Then
            Try
                serialPort.Write(StrConv(Chr(&H56), VbStrConv.None))
                serialPort.Write(StrConv(Chr(data1), VbStrConv.None))
                serialPort.Write(StrConv(Chr(data2), VbStrConv.None))
                serialPort.Write(StrConv(Chr(data3), VbStrConv.None))
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try

        End If
    End Sub

    'servo 23
    Private Sub VScrollBar30_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles VScrollBar30.ValueChanged
        TextBox26.Text = VScrollBar30.Value
        servo23 = VScrollBar30.Value

        data1 = (servo23 >> 6) And &HFF
        data2 = servo23 And &H3F
        data3 = NumericUpDown26.Value

        If s_en = 1 Then
            Try
                serialPort.Write(StrConv(Chr(&H57), VbStrConv.None))
                serialPort.Write(StrConv(Chr(data1), VbStrConv.None))
                serialPort.Write(StrConv(Chr(data2), VbStrConv.None))
                serialPort.Write(StrConv(Chr(data3), VbStrConv.None))
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try

        End If
    End Sub
    'servo 24

    Private Sub VScrollBar32_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles VScrollBar32.ValueChanged
        TextBox25.Text = VScrollBar32.Value
        servo22 = VScrollBar32.Value

        data1 = (servo24 >> 6) And &HFF
        data2 = servo24 And &H3F
        data3 = NumericUpDown25.Value

        If s_en = 1 Then
            Try
                serialPort.Write(StrConv(Chr(&H58), VbStrConv.None))
                serialPort.Write(StrConv(Chr(data1), VbStrConv.None))
                serialPort.Write(StrConv(Chr(data2), VbStrConv.None))
                serialPort.Write(StrConv(Chr(data3), VbStrConv.None))
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try

        End If
    End Sub
    'servo 25

    Private Sub VScrollBar31_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles VScrollBar31.ValueChanged
        TextBox24.Text = VScrollBar31.Value
        servo25 = VScrollBar31.Value

        data1 = (servo25 >> 6) And &HFF
        data2 = servo25 And &H3F
        data3 = NumericUpDown24.Value

        If s_en = 1 Then
            Try
                serialPort.Write(StrConv(Chr(&H59), VbStrConv.None))
                serialPort.Write(StrConv(Chr(data1), VbStrConv.None))
                serialPort.Write(StrConv(Chr(data2), VbStrConv.None))
                serialPort.Write(StrConv(Chr(data3), VbStrConv.None))
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try

        End If
    End Sub
    'servo 26

    Private Sub VScrollBar19_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles VScrollBar19.ValueChanged
        TextBox23.Text = VScrollBar19.Value
        servo26 = VScrollBar19.Value

        data1 = (servo26 >> 6) And &HFF
        data2 = servo26 And &H3F
        data3 = NumericUpDown23.Value

        If s_en = 1 Then
            Try
                serialPort.Write(StrConv(Chr(&H5A), VbStrConv.None))
                serialPort.Write(StrConv(Chr(data1), VbStrConv.None))
                serialPort.Write(StrConv(Chr(data2), VbStrConv.None))
                serialPort.Write(StrConv(Chr(data3), VbStrConv.None))
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try

        End If
    End Sub
    'servo 27

    Private Sub VScrollBar21_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles VScrollBar21.ValueChanged
        TextBox22.Text = VScrollBar21.Value
        servo27 = VScrollBar21.Value

        data1 = (servo27 >> 6) And &HFF
        data2 = servo27 And &H3F
        data3 = NumericUpDown22.Value

        If s_en = 1 Then
            Try
                serialPort.Write(StrConv(Chr(&H5B), VbStrConv.None))
                serialPort.Write(StrConv(Chr(data1), VbStrConv.None))
                serialPort.Write(StrConv(Chr(data2), VbStrConv.None))
                serialPort.Write(StrConv(Chr(data3), VbStrConv.None))
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try

        End If
    End Sub
    'servo 28

    Private Sub VScrollBar23_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles VScrollBar23.ValueChanged
        TextBox21.Text = VScrollBar23.Value
        servo28 = VScrollBar23.Value

        data1 = (servo28 >> 6) And &HFF
        data2 = servo28 And &H3F
        data3 = NumericUpDown21.Value

        If s_en = 1 Then
            Try
                serialPort.Write(StrConv(Chr(&H5C), VbStrConv.None))
                serialPort.Write(StrConv(Chr(data1), VbStrConv.None))
                serialPort.Write(StrConv(Chr(data2), VbStrConv.None))
                serialPort.Write(StrConv(Chr(data3), VbStrConv.None))
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try

        End If
    End Sub
    'servo 29

    Private Sub VScrollBar25_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles VScrollBar25.ValueChanged
        TextBox20.Text = VScrollBar25.Value
        servo29 = VScrollBar25.Value

        data1 = (servo29 >> 6) And &HFF
        data2 = servo29 And &H3F
        data3 = NumericUpDown20.Value

        If s_en = 1 Then
            Try
                serialPort.Write(StrConv(Chr(&H5D), VbStrConv.None))
                serialPort.Write(StrConv(Chr(data1), VbStrConv.None))
                serialPort.Write(StrConv(Chr(data2), VbStrConv.None))
                serialPort.Write(StrConv(Chr(data3), VbStrConv.None))
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try

        End If
    End Sub
    'servo 30 

    Private Sub VScrollBar27_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles VScrollBar27.ValueChanged
        TextBox19.Text = VScrollBar27.Value
        servo30 = VScrollBar27.Value

        data1 = (servo30 >> 6) And &HFF
        data2 = servo30 And &H3F
        data3 = NumericUpDown19.Value

        If s_en = 1 Then
            Try
                serialPort.Write(StrConv(Chr(&H5E), VbStrConv.None))
                serialPort.Write(StrConv(Chr(data1), VbStrConv.None))
                serialPort.Write(StrConv(Chr(data2), VbStrConv.None))
                serialPort.Write(StrConv(Chr(data3), VbStrConv.None))
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try

        End If
    End Sub

    'servo 31

    Private Sub VScrollBar29_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles VScrollBar29.ValueChanged
        TextBox18.Text = VScrollBar29.Value
        servo31 = VScrollBar29.Value
        data3 = NumericUpDown18.Value

        data1 = (servo31 >> 6) And &HFF
        data2 = servo31 And &H3F

        If s_en = 1 Then
            Try
                serialPort.Write(StrConv(Chr(&H5F), VbStrConv.None))
                serialPort.Write(StrConv(Chr(data1), VbStrConv.None))
                serialPort.Write(StrConv(Chr(data2), VbStrConv.None))
                serialPort.Write(StrConv(Chr(data3), VbStrConv.None))
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try

        End If
    End Sub
    'servo 32

    Private Sub VScrollBar18_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles VScrollBar18.ValueChanged
        TextBox17.Text = VScrollBar18.Value
        servo32 = VScrollBar18.Value

        data1 = (servo32 >> 6) And &HFF
        data2 = servo32 And &H3F
        data3 = NumericUpDown17.Value

        If s_en = 1 Then
            Try
                serialPort.Write(StrConv(Chr(&H60), VbStrConv.None))
                serialPort.Write(StrConv(Chr(data1), VbStrConv.None))
                serialPort.Write(StrConv(Chr(data2), VbStrConv.None))
                serialPort.Write(StrConv(Chr(data3), VbStrConv.None))

            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try

        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button2.Click
        VScrollBar17.Value = 731
        VScrollBar18.Value = 731
        VScrollBar19.Value = 731
        VScrollBar20.Value = 731
        VScrollBar21.Value = 731
        VScrollBar22.Value = 731
        VScrollBar23.Value = 731
        VScrollBar24.Value = 731
        VScrollBar25.Value = 731
        VScrollBar26.Value = 731
        VScrollBar27.Value = 731
        VScrollBar28.Value = 731
        VScrollBar29.Value = 731
        VScrollBar30.Value = 731
        VScrollBar31.Value = 731
        VScrollBar32.Value = 731
        NumericUpDown17.Value = 0
        NumericUpDown18.Value = 0
        NumericUpDown19.Value = 0
        NumericUpDown20.Value = 0
        NumericUpDown21.Value = 0
        NumericUpDown22.Value = 0
        NumericUpDown23.Value = 0
        NumericUpDown24.Value = 0
        NumericUpDown25.Value = 0
        NumericUpDown26.Value = 0
        NumericUpDown27.Value = 0
        NumericUpDown28.Value = 0
        NumericUpDown29.Value = 0
        NumericUpDown30.Value = 0
        NumericUpDown31.Value = 0
        NumericUpDown32.Value = 0
    End Sub


    Private Sub LinkLabel1_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
      
        System.Diagnostics.Process.Start("www.cytron.com.my")
    End Sub

    
End Class


