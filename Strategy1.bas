Attribute VB_Name = "Strategy1"
Rem: KEYSTRAT 1.7 - Strategy1.bas
Rem: ****************************
Option Explicit
Rem: Declare Variables
Dim s1_stage As Integer
Dim s1_last As Integer
Dim s1_current As Integer
Rem: Control Strategy Stages
Public Sub Strategy1_Move()
    If s1_stage = 0 Then
        Strategy1_Move_Stage1
    Else
        Main.Keyboard_Array_Click (s1_last)
        s1_stage = 0
        s1_current = 0
    End If
End Sub
Rem: Control Strategy Stage1
Private Sub Strategy1_Move_Stage1()
        Main.Keyboard_Array(s1_last).BackColor = &H80000016
        Main.Keyboard_Array(s1_current).BackColor = &HFFFF&
        s1_last = s1_current
        s1_current = (s1_current + 1) Mod 56
End Sub
Rem: Increment Stage Counter
Public Sub Strategy1_Inc_Stage()
    s1_stage = s1_stage + 1
    Main.Clear_Keyboard
End Sub
Rem: Reset Variables
Public Sub Strategy1_Reset()
    s1_stage = 0
    s1_last = 0
    s1_current = 0
End Sub

