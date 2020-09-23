Attribute VB_Name = "Strategy3"
Rem: KEYSTRAT 1.7 - Strategy3.bas
Rem: ****************************
Option Explicit
Rem: Declare Variables
Dim s3_stage As Integer
Dim s3_last As Integer
Dim s3_current As Integer
Dim s3_i As Integer
Dim s3_offset As Integer
Dim s3_position As Integer
Rem: Control Strategy Stages
Public Sub Strategy3_Move()
    If s3_stage = 0 Then
        Strategy3_Move_Stage1
    ElseIf s3_stage = 1 Then
        Strategy3_Move_Stage2
    Else
        Main.Keyboard_Array_Click (s3_position)
        Strategy3_Reset
    End If
End Sub
Rem: Control Strategy Stage1
Private Sub Strategy3_Move_Stage1()
    For s3_i = 0 To 6
        Main.Keyboard_Array(s3_last + s3_i).BackColor = &H80000016
    Next s3_i
    For s3_i = 0 To 6
        Main.Keyboard_Array(s3_current + s3_i).BackColor = &HFFFF&
    Next s3_i
    s3_last = s3_current
    s3_current = (s3_current + 7) Mod 56
End Sub
Rem: Control Strategy Stage2
Private Sub Strategy3_Move_Stage2()
    Main.Keyboard_Array(s3_last + s3_offset).BackColor = &H80000016
    Main.Keyboard_Array(s3_current + s3_offset).BackColor = &HFFFF&
    s3_last = s3_current
    s3_current = (s3_current + 1) Mod 7
End Sub
Rem: Increment Stage Counter
Public Sub Strategy3_Inc_Stage()
    s3_stage = s3_stage + 1
    s3_position = s3_last + s3_offset
    s3_offset = s3_last
    s3_current = 0
    s3_last = 0
    Main.Clear_Keyboard
End Sub
Rem: Reset Variables
Public Sub Strategy3_Reset()
    s3_stage = 0
    s3_last = 0
    s3_current = 0
    s3_i = 0
    s3_offset = 0
    s3_position = 0
End Sub

