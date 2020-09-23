Attribute VB_Name = "Strategy2"
Rem: KEYSTRAT 1.7 - Strategy2.bas
Rem: ****************************
Option Explicit
Rem: Declare Variables
Dim s2_stage As Integer
Dim s2_last As Integer
Dim s2_current As Integer
Dim s2_i As Integer
Dim s2_offset As Integer
Dim s2_position As Integer
Rem: Control Strategy Stages
Public Sub Strategy2_Move()
    If s2_stage = 0 Then
        Strategy2_Move_Stage1
    ElseIf s2_stage = 1 Then
        Strategy2_Move_Stage2
    Else
        Main.Keyboard_Array_Click (s2_position)
        Strategy2_Reset
    End If
End Sub
Rem: Control Strategy Stage1
Private Sub Strategy2_Move_Stage1()
    For s2_i = 0 To 3
        Main.Keyboard_Array(s2_last + s2_i).BackColor = &H80000016
    Next s2_i
    For s2_i = 0 To 3
        Main.Keyboard_Array(s2_current + s2_i).BackColor = &HFFFF&
    Next s2_i
    s2_last = s2_current
    s2_current = (s2_current + 4) Mod 56
End Sub
Rem: Control Strategy Stage2
Private Sub Strategy2_Move_Stage2()
    Main.Keyboard_Array(s2_last + s2_offset).BackColor = &H80000016
    Main.Keyboard_Array(s2_current + s2_offset).BackColor = &HFFFF&
    s2_last = s2_current
    s2_current = (s2_current + 1) Mod 4
End Sub
Rem: Increment Stage Counter
Public Sub Strategy2_Inc_Stage()
    s2_stage = s2_stage + 1
    s2_position = s2_last + s2_offset
    s2_offset = s2_last
    s2_current = 0
    s2_last = 0
    Main.Clear_Keyboard
End Sub
Rem: Reset Variables
Public Sub Strategy2_Reset()
    s2_stage = 0
    s2_last = 0
    s2_current = 0
    s2_i = 0
    s2_offset = 0
    s2_position = 0
End Sub
