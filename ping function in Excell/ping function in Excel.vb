Function Ping(strip)
Dim objshell, boolcode
Set objshell = CreateObject("Wscript.Shell")
boolcode = objshell.Run("ping -n 1 -w 1000 " & strip, 0, True)
If boolcode = 0 Then
    Ping = True
Else
    Ping = False
End If
End Function
Sub PingSystem()
    Dim strip As String
    Dim test As String
    Dim Online, Offline, intName, intPing As Integer
    Online = 0
    Offline = 0
    intID = 2
    intName = 3
    intPing = 5
    intOn = 10
    intOff = 11
    ActiveSheet.Range("F1").Value = "RUNNING"
    ActiveSheet.Cells(1, intOn).Font.Color = RGB(146, 208, 80)
    ActiveSheet.Cells(1, intOff).Font.Color = RGB(200, 0, 0)
    ActiveSheet.Cells(1, intOn).Value = "Online - " & Online
    ActiveSheet.Cells(1, intOff).Value = "Offline - " & Offline
    ActiveSheet.Cells(1, 12).Value = ""
    Do
        For introw = 3 To ActiveSheet.Cells(65536, 4).End(xlUp).Row + 1
            If IsEmpty(ActiveSheet.Cells(introw, 4).Value) Then
                ActiveSheet.Range("F1").Value = "Successfuly"
                Exit For
            Else
                strip = ActiveSheet.Cells(introw, 4).Value
                If Ping(strip) = True Then
                    ActiveSheet.Cells(introw, intName).Interior.Color = RGB(146, 208, 80)
                    ActiveSheet.Cells(introw, intPing).Interior.Color = RGB(146, 208, 80)
                    ActiveSheet.Cells(introw, intPing).Value = "Online"
                    'Application.Wait (Now + TimeValue("0:00:01"))
                    ActiveSheet.Cells(introw, intName).Font.Color = RGB(0, 0, 0)
                    ActiveSheet.Cells(introw, intPing).Font.Color = RGB(0, 0, 0)
                    If ActiveSheet.Cells(introw, intID).Value <> "-" Then
                       Online = Online + 1
                       ActiveSheet.Cells(1, intOn).Value = "Online - " & Online
                    End If
                    If introw = 3 Then
                        ActiveSheet.Cells(introw, intName).Interior.Color = RGB(0, 176, 80)
                        ActiveSheet.Cells(introw, intPing).Interior.Color = RGB(0, 176, 80)
                    End If
                Else
                    ActiveSheet.Cells(introw, intName).Font.Color = RGB(200, 0, 0)
                    ActiveSheet.Cells(introw, intPing).Font.Color = RGB(200, 0, 0)
                    ActiveSheet.Cells(introw, intPing).Value = "Offline"
                    'Application.Wait (Now + TimeValue("0:00:01"))
                    ActiveSheet.Cells(introw, intName).Interior.Color = RGB(255, 255, 0)
                    ActiveSheet.Cells(introw, intPing).Interior.Color = RGB(255, 255, 0)
                    If ActiveSheet.Cells(introw, intID).Value <> "-" Then
                       Offline = Offline + 1
                       ActiveSheet.Cells(1, intOff).Value = "Offline - " & Offline
                    End If
                End If
            End If
            If ActiveSheet.Range("F1").Value = "STOP" Then
                Exit Do
            End If
        Next
        ActiveSheet.Cells(1, 12).Value = "Last Check: " & Date & " " & Time()
    Loop Until ActiveSheet.Range("F1").Value = "Successfuly"
End Sub

Sub stop_ping()
    ActiveSheet.Range("F1").Value = "STOP"
End Sub

Sub add_new_line()
    For introw = 3 To ActiveSheet.Cells(65536, 2).End(xlUp).Row + 1
        If IsEmpty(ActiveSheet.Cells(introw, 2).Value) Then
            ' Если цвет ячейки уже заполнен, то выходим из цикла
            If ActiveSheet.Cells(introw, 2).Interior.Color = 5296274 Then
                Exit For
            End If
            ' Копирование диапазона выше в новую строку с дальнейшей очисткой
            ActiveSheet.Range(ActiveSheet.Cells(introw - 1, 2), ActiveSheet.Cells(introw - 1, 14)).Select
            Selection.Copy
            ActiveSheet.Range(ActiveSheet.Cells(introw, 2), ActiveSheet.Cells(introw, 14)).Select
            ActiveSheet.Paste
            Selection.ClearContents
            With Selection.Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .Color = 5296274
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
            Range("A2").Select

        End If
    Next
End Sub
