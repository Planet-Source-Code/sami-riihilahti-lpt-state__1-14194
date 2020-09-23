Attribute VB_Name = "Module1"
'With Sleep function, system uses only minimal CPU-usage
Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
'With Inp32, we can read ports fast with visual basic
Public Declare Function Inp Lib "inpout32.dll" _
    Alias "Inp32" (ByVal PortAddress As Integer) As Integer
'We could use Out32 to write to ports, but we don't need it
'Public Declare Sub Out Lib "inpout32.dll" _
    Alias "Out32" (ByVal PortAddress As Integer, ByVal Value As Integer)

Public DataLeds As Integer 'Integer to store port input
Public UserCancel As Boolean 'We use it to end program loops properly after user has pressed exit
Public LPTport As Integer 'Current port address being used
Public ByteExt As Integer 'Loop variable
Public SleepValue As Integer 'Interval for updating 'leds'
Sub Main() 'The main sub, which starts first
    LPTport = 888 'Set default port address (LPT1)
    SleepValue = 1 'Set default update interval
    Form1.Show 'Show main form
    Do
        DataLeds = Inp(LPTport) 'Data outputs 0-7
        For ByteExt = 7 To 0 Step -1
            If DataLeds And 2 ^ ByteExt Then
                Set Form1.LPTdataled(ByteExt).Picture = Form1.led_green.Picture
            Else:
                Set Form1.LPTdataled(ByteExt).Picture = Form1.led_gray.Picture
            End If
        Next
        
        DataLeds = Inp(LPTport + 1) 'LPT port Feedbacks bytes 3 to 7
        If DataLeds And 8 Then
            Form1.LPTerrorled.Picture = Form1.led_yellow.Picture
        Else:
            Form1.LPTerrorled.Picture = Form1.led_gray.Picture
        End If
        If DataLeds And 16 Then
            Form1.LPTselectled.Picture = Form1.led_yellow.Picture
        Else:
            Form1.LPTselectled.Picture = Form1.led_gray.Picture
        End If
        If DataLeds And 32 Then
            Form1.LPTpaperled.Picture = Form1.led_yellow.Picture
        Else:
            Form1.LPTpaperled.Picture = Form1.led_gray.Picture
        End If
        If DataLeds And 64 Then
            Form1.LPTackled.Picture = Form1.led_yellow.Picture
        Else:
            Form1.LPTackled.Picture = Form1.led_gray.Picture
        End If
        If DataLeds And 128 Then
            Form1.LPTbusyled.Picture = Form1.led_yellow.Picture
        Else:
            Form1.LPTbusyled.Picture = Form1.led_gray.Picture
        End If
        
        DataLeds = Inp(LPTport + 2) 'LPT port Controls bytes 0 to 3
        If DataLeds And 1 Then
            Form1.LPTstrobeled.Picture = Form1.led_green.Picture
        Else:
            Form1.LPTstrobeled.Picture = Form1.led_gray.Picture
        End If
        If DataLeds And 2 Then
            Form1.LPTautofeedled.Picture = Form1.led_green.Picture
        Else:
            Form1.LPTautofeedled.Picture = Form1.led_gray.Picture
        End If
        If DataLeds And 4 Then
            Form1.LPTinitled.Picture = Form1.led_green.Picture
        Else:
            Form1.LPTinitled.Picture = Form1.led_gray.Picture
        End If
        If DataLeds And 8 Then
            Form1.LPTselectinled.Picture = Form1.led_green.Picture
        Else:
            Form1.LPTselectinled.Picture = Form1.led_gray.Picture
        End If
        
        Sleep SleepValue 'Pause this application in order to keep CPU-usage low
        DoEvents 'Update form
    Loop Until UserCancel = True 'Loop exits when user exits this application
End
End Sub
