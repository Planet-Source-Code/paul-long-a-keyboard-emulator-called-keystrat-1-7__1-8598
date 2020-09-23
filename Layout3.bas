Attribute VB_Name = "Layout3"
Rem: KEYSTRAT 1.7 - Layout3.bas (Alphabetic)
Rem: ***************************************
Option Explicit
Rem: Declare Variables
Dim layout3_lockon As Boolean
Dim layout3_shifton As Boolean
Rem: Setup Layout According to Lock and Shift Status
Public Sub Layout3_Setup()
    If (Main.Shift_Light.BackColor = &HFF00&) Then 'Shift Light On
        Layout3_Shift_On
        layout3_shifton = True
    Else 'Shift Light Off
        Layout3_Shift_Off
        layout3_shifton = False
    End If
    If (Main.Caps_Light.BackColor = &HFF00&) Then 'Lock Light On
        Layout3_Lock_On
        layout3_lockon = True
    Else 'Lock Light Off
        Layout3_Lock_Off
        layout3_lockon = False
    End If
    Layout3_Static_Keys
End Sub
Rem: Toggle Lock Status
Public Sub Layout3_Toggle_Lock()
    If layout3_lockon Then
        Layout3_Lock_Off
        layout3_lockon = False
    Else
        Layout3_Lock_On
        layout3_lockon = True
    End If
End Sub
Rem: Toggle Shift Status
Public Sub Layout3_Toggle_Shift()
    If layout3_shifton Then
        Layout3_Shift_Off
        layout3_shifton = False
    Else
        Layout3_Shift_On
        layout3_shifton = True
    End If
End Sub
Rem: Static Keys
Private Sub Layout3_Static_Keys()
    Main.Keyboard_Array(13).Caption = "bksp"  ' Label key with bksp
    Main.Keyboard_Array(14).Caption = "tab"   ' Label key with tab
    Main.Keyboard_Array(28).Caption = "lock"  ' Label key with lock
    Main.Keyboard_Array(40).Caption = "enter" ' Label key with enter
    Main.Keyboard_Array(41).Caption = "shift" ' Label key with shift
    Main.Keyboard_Array(53).Caption = "clear" ' Label key with clear
    Main.Keyboard_Array(54).Caption = "space" ' Label key with space
    Main.Keyboard_Array(55).Caption = "stop"  ' Label key with clear
End Sub
Rem: Lock On
Private Sub Layout3_Lock_On()
    Main.Keyboard_Array(15).Caption = "A"     ' Relabel key
    Main.Keyboard_Array(16).Caption = "B"     ' Relabel key
    Main.Keyboard_Array(17).Caption = "C"     ' Relabel key
    Main.Keyboard_Array(18).Caption = "D"     ' Relabel key
    Main.Keyboard_Array(19).Caption = "E"     ' Relabel key
    Main.Keyboard_Array(20).Caption = "F"     ' Relabel key
    Main.Keyboard_Array(21).Caption = "G"     ' Relabel key
    Main.Keyboard_Array(22).Caption = "H"     ' Relabel key
    Main.Keyboard_Array(23).Caption = "I"     ' Relabel key
    Main.Keyboard_Array(24).Caption = "J"     ' Relabel key
    Main.Keyboard_Array(29).Caption = "K"     ' Relabel key
    Main.Keyboard_Array(30).Caption = "L"     ' Relabel key
    Main.Keyboard_Array(31).Caption = "M"     ' Relabel key
    Main.Keyboard_Array(32).Caption = "N"     ' Relabel key
    Main.Keyboard_Array(33).Caption = "O"     ' Relabel key
    Main.Keyboard_Array(34).Caption = "P"     ' Relabel key
    Main.Keyboard_Array(35).Caption = "Q"     ' Relabel key
    Main.Keyboard_Array(36).Caption = "R"     ' Relabel key
    Main.Keyboard_Array(37).Caption = "S"     ' Relabel key
    Main.Keyboard_Array(43).Caption = "T"     ' Relabel key
    Main.Keyboard_Array(44).Caption = "U"     ' Relabel key
    Main.Keyboard_Array(45).Caption = "V"     ' Relabel key
    Main.Keyboard_Array(46).Caption = "W"     ' Relabel key
    Main.Keyboard_Array(47).Caption = "X"     ' Relabel key
    Main.Keyboard_Array(48).Caption = "Y"     ' Relabel key
    Main.Keyboard_Array(49).Caption = "Z"     ' Relabel key
End Sub
Rem: Lock Off
Private Sub Layout3_Lock_Off()
    Main.Keyboard_Array(15).Caption = "a"     ' Relabel key
    Main.Keyboard_Array(16).Caption = "b"     ' Relabel key
    Main.Keyboard_Array(17).Caption = "c"     ' Relabel key
    Main.Keyboard_Array(18).Caption = "d"     ' Relabel key
    Main.Keyboard_Array(19).Caption = "e"     ' Relabel key
    Main.Keyboard_Array(20).Caption = "f"     ' Relabel key
    Main.Keyboard_Array(21).Caption = "g"     ' Relabel key
    Main.Keyboard_Array(22).Caption = "h"     ' Relabel key
    Main.Keyboard_Array(23).Caption = "i"     ' Relabel key
    Main.Keyboard_Array(24).Caption = "j"     ' Relabel key
    Main.Keyboard_Array(29).Caption = "k"     ' Relabel key
    Main.Keyboard_Array(30).Caption = "l"     ' Relabel key
    Main.Keyboard_Array(31).Caption = "m"     ' Relabel key
    Main.Keyboard_Array(32).Caption = "n"     ' Relabel key
    Main.Keyboard_Array(33).Caption = "o"     ' Relabel key
    Main.Keyboard_Array(34).Caption = "p"     ' Relabel key
    Main.Keyboard_Array(35).Caption = "q"     ' Relabel key
    Main.Keyboard_Array(36).Caption = "r"     ' Relabel key
    Main.Keyboard_Array(37).Caption = "s"     ' Relabel key
    Main.Keyboard_Array(43).Caption = "t"     ' Relabel key
    Main.Keyboard_Array(44).Caption = "u"     ' Relabel key
    Main.Keyboard_Array(45).Caption = "v"     ' Relabel key
    Main.Keyboard_Array(46).Caption = "w"     ' Relabel key
    Main.Keyboard_Array(47).Caption = "x"     ' Relabel key
    Main.Keyboard_Array(48).Caption = "y"     ' Relabel key
    Main.Keyboard_Array(49).Caption = "z"     ' Relabel key
End Sub
Rem: Shift On
Private Sub Layout3_Shift_On()
    Main.Keyboard_Array(0).Caption = "¬"      ' Relabel key
    Main.Keyboard_Array(1).Caption = "!"      ' Relabel key
    Main.Keyboard_Array(2).Caption = Chr(34)  ' Relabel key (with ")
    Main.Keyboard_Array(3).Caption = "£"      ' Relabel key
    Main.Keyboard_Array(4).Caption = "$"      ' Relabel key
    Main.Keyboard_Array(5).Caption = "%"      ' Relabel key
    Main.Keyboard_Array(6).Caption = "^"      ' Relabel key
    Main.Keyboard_Array(7).Caption = "&&"     ' Have to use && to display & on key
    Main.Keyboard_Array(8).Caption = "*"      ' Relabel key
    Main.Keyboard_Array(9).Caption = "("      ' Relabel key
    Main.Keyboard_Array(10).Caption = ")"     ' Relabel key
    Main.Keyboard_Array(11).Caption = "_"     ' Relabel key
    Main.Keyboard_Array(12).Caption = "+"     ' Relabel key
    Main.Keyboard_Array(25).Caption = "{"     ' Relabel key
    Main.Keyboard_Array(26).Caption = "}"     ' Relabel key
    Main.Keyboard_Array(27).Caption = "~"     ' Relabel key
    Main.Keyboard_Array(38).Caption = ":"     ' Relabel key
    Main.Keyboard_Array(39).Caption = "@"     ' Relabel key
    Main.Keyboard_Array(42).Caption = "|"     ' Relabel key
    Main.Keyboard_Array(50).Caption = "<"     ' Relabel key
    Main.Keyboard_Array(51).Caption = ">"     ' Relabel key
    Main.Keyboard_Array(52).Caption = "?"     ' Relabel key
End Sub
Rem: Shift Off
Public Sub Layout3_Shift_Off()
    Main.Keyboard_Array(0).Caption = "`"      ' Relabel key
    Main.Keyboard_Array(1).Caption = "1"      ' Relabel key
    Main.Keyboard_Array(2).Caption = "2"      ' Relabel key
    Main.Keyboard_Array(3).Caption = "3"      ' Relabel key
    Main.Keyboard_Array(4).Caption = "4"      ' Relabel key
    Main.Keyboard_Array(5).Caption = "5"      ' Relabel key
    Main.Keyboard_Array(6).Caption = "6"      ' Relabel key
    Main.Keyboard_Array(7).Caption = "7"      ' Relabel key
    Main.Keyboard_Array(8).Caption = "8"      ' Relabel key
    Main.Keyboard_Array(9).Caption = "9"      ' Relabel key
    Main.Keyboard_Array(10).Caption = "0"     ' Relabel key
    Main.Keyboard_Array(11).Caption = "-"     ' Relabel key
    Main.Keyboard_Array(12).Caption = "="     ' Relabel key
    Main.Keyboard_Array(25).Caption = "["     ' Relabel key
    Main.Keyboard_Array(26).Caption = "]"     ' Relabel key
    Main.Keyboard_Array(27).Caption = "#"     ' Relabel key
    Main.Keyboard_Array(38).Caption = ";"     ' Relabel key
    Main.Keyboard_Array(39).Caption = "'"     ' Relabel key
    Main.Keyboard_Array(42).Caption = "\"     ' Relabel key
    Main.Keyboard_Array(50).Caption = ","     ' Relabel key
    Main.Keyboard_Array(51).Caption = "."     ' Relabel key
    Main.Keyboard_Array(52).Caption = "/"     ' Relabel key
End Sub

