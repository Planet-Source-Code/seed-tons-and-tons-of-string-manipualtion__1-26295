Attribute VB_Name = "LongFuncs"
Public Function Elite(Txt As String)
made$ = ""
For q = 1 To Len(Txt)
    letter$ = ""
    letter$ = Mid$(Txt, q, 1)
    Leet$ = ""
    X = Int(Rnd * 3 + 1)
    If letter$ = "a" Then
    If X = 1 Then Leet$ = "�"
    If X = 2 Then Leet$ = "�"
    If X = 3 Then Leet$ = "�"
    End If
    If letter$ = "b" Then Leet$ = "b"
    If letter$ = "c" Then Leet$ = "�"
    If letter$ = "e" Then
    If X = 1 Then Leet$ = "�"
    If X = 2 Then Leet$ = "�"
    If X = 3 Then Leet$ = "�"
    End If
    If letter$ = "i" Then
    If X = 1 Then Leet$ = "�"
    If X = 2 Then Leet$ = "�"
    If X = 3 Then Leet$ = "�"
    End If
    If letter$ = "j" Then Leet$ = ",j"
    If letter$ = "n" Then Leet$ = "�"
    If letter$ = "o" Then
    If X = 1 Then Leet$ = "�"
    If X = 2 Then Leet$ = "�"
    If X = 3 Then Leet$ = "�"
    End If
    If letter$ = "s" Then Leet$ = "�"
    If letter$ = "t" Then Leet$ = "�"
    If letter$ = "u" Then
    If X = 1 Then Leet$ = "�"
    If X = 2 Then Leet$ = "�"
    If X = 3 Then Leet$ = "�"
    End If
    If letter$ = "w" Then Leet$ = "vv"
    If letter$ = "y" Then Leet$ = "�"
    If letter$ = "0" Then Leet$ = "�"
    If letter$ = "A" Then
    If X = 1 Then Leet$ = "�"
    If X = 2 Then Leet$ = "�"
    If X = 3 Then Leet$ = "�"
    End If
    If letter$ = "B" Then Leet$ = "�"
    If letter$ = "C" Then Leet$ = "�"
    If letter$ = "D" Then Leet$ = "�"
    If letter$ = "E" Then Leet$ = "�"
    If letter$ = "I" Then
    If X = 1 Then Leet$ = "�"
    If X = 2 Then Leet$ = "�"
    If X = 3 Then Leet$ = "�"
    End If
    If letter$ = "N" Then Leet$ = "�"
    If letter$ = "O" Then Leet$ = "�"
    If letter$ = "S" Then Leet$ = "�"
    If letter$ = "U" Then Leet$ = "�"
    If letter$ = "W" Then Leet$ = "VV"
    If letter$ = "Y" Then Leet$ = "�"
    If Len(Leet$) = 0 Then Leet$ = letter$
    made$ = made$ & Leet$
Next q
Txt = made$
frmMain.txtBox.text = Txt
End Function
