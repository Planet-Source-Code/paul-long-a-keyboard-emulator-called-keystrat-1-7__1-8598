Attribute VB_Name = "Strategy4"
Rem: KEYSTRAT 1.7 - Strategy4.bas
Rem: ****************************
Option Explicit
Rem: Declare Variables
Dim s4_stage As Integer
Dim s4_row As Integer
Dim s4_last_start As Integer
Dim s4_last_end As Integer
Dim s4_current_start As Integer
Dim s4_current_end As Integer
Dim s4_i As Integer
Dim s4_length As Integer
Dim s4_last As Integer
Dim s4_current As Integer
Rem: Control Strategy Stages
Public Sub Strategy4_Move()
    If s4_stage = 0 Then
        Strategy4_Move_Stage1
    ElseIf s4_stage = 1 Then
        Strategy4_Move_Stage2
    Else
        Main.Keyboard_Array_Click (s4_last + s4_current_start)
        Strategy4_Reset
    End If
End Sub
Rem: Control Strategy Stage1
Private Sub Strategy4_Move_Stage1()
    If s4_row = 0 Then
        s4_last_start = 54
        s4_last_end = 55
        s4_current_start = 0
        s4_current_end = 13
    ElseIf s4_row = 1 Then
        s4_last_start = s4_current_start
        s4_last_end = s4_current_end
        s4_current_start = 14
        s4_current_end = 27
    ElseIf s4_row = 2 Then
        s4_last_start = s4_current_start
        s4_last_end = s4_current_end
        s4_current_start = 28
        s4_current_end = 40
    ElseIf s4_row = 3 Then
        s4_last_start = s4_current_start
        s4_last_end = s4_current_end
        s4_current_start = 41
        s4_current_end = 53
    ElseIf s4_row = 4 Then
        s4_last_start = s4_current_start
        s4_last_end = s4_current_end
        s4_current_start = 54
        s4_current_end = 55
    End If
    For s4_i = s4_last_start To s4_last_end
        Main.Keyboard_Array(s4_i).BackColor = &H80000016
    Next s4_i
    For s4_i = s4_current_start To s4_current_end
        Main.Keyboard_Array(s4_i).BackColor = &HFFFF&
    Next s4_i
    s4_row = (s4_row + 1) Mod 5
End Sub
Rem: Control Strategy Stage2
Private Sub Strategy4_Move_Stage2()
        s4_length = (s4_current_end - s4_current_start) + 1
        Main.Keyboard_Array(s4_last + s4_current_start).BackColor = &H80000016
        Main.Keyboard_Array(s4_current + s4_current_start).BackColor = &HFFFF&
        s4_last = s4_current
        s4_current = (s4_current + 1) Mod s4_length
End Sub
Rem: Increment Stage Counter
Public Sub Strategy4_Inc_Stage()
    s4_stage = s4_stage + 1
    Main.Clear_Keyboard
End Sub
Rem: Reset Variables
Public Sub Strategy4_Reset()
    s4_stage = 0
    s4_row = 0
    s4_last_start = 0
    s4_last_end = 0
    s4_current_start = 0
    s4_current_end = 0
    s4_i = 0
    s4_length = 0
    s4_last = 0
    s4_current = 0
End Sub
