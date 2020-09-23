VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Crazy String Manipulation!"
   ClientHeight    =   3225
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4815
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3225
   ScaleWidth      =   4815
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdBI 
      Caption         =   "Bold/Italics"
      Height          =   255
      Left            =   3000
      TabIndex        =   25
      Top             =   2880
      Width           =   1215
   End
   Begin VB.CommandButton cmdULine 
      Caption         =   "UnderLine"
      Height          =   255
      Left            =   3000
      TabIndex        =   24
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton cmdStrike 
      Caption         =   "Strike"
      Height          =   255
      Left            =   3000
      TabIndex        =   23
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton cmdLamer 
      Caption         =   "lAmEr"
      Height          =   255
      Left            =   3000
      TabIndex        =   22
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton cmdDouble 
      Appearance      =   0  'Flat
      Caption         =   "Double"
      Height          =   255
      Left            =   3000
      TabIndex        =   21
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton cmdSwap 
      Caption         =   "Swap"
      Height          =   255
      Left            =   3000
      TabIndex        =   20
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton cmdRepOne 
      Caption         =   "Exchange 1"
      Height          =   255
      Left            =   3000
      TabIndex        =   19
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton cmdExchange 
      Caption         =   "Cut Char(s)"
      Height          =   255
      Left            =   1800
      TabIndex        =   18
      Top             =   2880
      Width           =   1215
   End
   Begin VB.CommandButton cmdReplace 
      Caption         =   "Replace"
      Height          =   255
      Left            =   1800
      TabIndex        =   17
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton cmdLTrim 
      Caption         =   "Left Trim"
      Height          =   255
      Left            =   1800
      TabIndex        =   16
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton cmdCutSpaces 
      Caption         =   "Right Trim"
      Height          =   255
      Left            =   1800
      TabIndex        =   15
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton cmdTypy 
      Caption         =   "Typy Text"
      Height          =   255
      Left            =   1800
      TabIndex        =   14
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton cmdCool 
      Caption         =   "Color Me"
      Height          =   255
      Left            =   1800
      TabIndex        =   13
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton cmdDoubleRev 
      Caption         =   "Double Rev."
      Height          =   255
      Left            =   1800
      TabIndex        =   12
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton cmdRevSen 
      Caption         =   "Rev. Sentence"
      Height          =   255
      Left            =   600
      TabIndex        =   11
      Top             =   2880
      Width           =   1215
   End
   Begin VB.CommandButton cmdJumble 
      Caption         =   "Jumble"
      Height          =   255
      Left            =   600
      TabIndex        =   10
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton cmdElite 
      Caption         =   """Elite"""
      Height          =   255
      Left            =   600
      TabIndex        =   9
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton cmdHacker 
      Caption         =   "HaCKeR"
      Height          =   255
      Left            =   600
      TabIndex        =   8
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear A&ll"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1080
      Width           =   4575
   End
   Begin VB.CommandButton cmdIns 
      Caption         =   "Insert Special"
      Height          =   255
      Left            =   600
      TabIndex        =   6
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton cmdSpace 
      Caption         =   "Space Out"
      Height          =   255
      Left            =   600
      TabIndex        =   5
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "<"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   4320
      TabIndex        =   4
      Top             =   800
      Width           =   375
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   ">"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   4320
      TabIndex        =   3
      Top             =   600
      Width           =   375
   End
   Begin VB.CommandButton cmdRev 
      Caption         =   "&Reverse"
      Height          =   255
      Left            =   600
      TabIndex        =   2
      Top             =   1440
      Width           =   1215
   End
   Begin VB.TextBox txtBox 
      Height          =   405
      Left            =   120
      MaxLength       =   200
      TabIndex        =   1
      Text            =   "Type in here!"
      Top             =   600
      Width           =   4095
   End
   Begin VB.CommandButton SCR 
      Caption         =   "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4575
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Saved$

Sub Pause(Time)
Current = Timer
Do While Timer - Current < Time
DoEvents
Loop
End Sub

Private Sub cmdBI_Click()
If txtBox.Font.Bold = False Then txtBox.Font.Bold = True: Exit Sub
If txtBox.Font.Italic = False Then txtBox.Font.Italic = True: Exit Sub
If txtBox.Font.Bold = False Then
txtBox.Font.Bold = False
txtBox.Font.Italic = True
Else
txtBox.Font.Bold = False
txtBox.Font.Italic = False
End If
txtBox.SetFocus
End Sub

Private Sub cmdClear_Click()
txtBox = "": txtBox.SetFocus
End Sub

Private Sub cmdCool_Click()
Dim in1, in2 As Integer
in1 = Int(Rnd * 256)
in2 = Int(Rnd * 256)
For e = 0 To 255
Me.BackColor = RGB(in1, in1, in2)
Me.BackColor = RGB(in2, in1, in2)
Me.BackColor = RGB(in2, in1, in2)
Me.BackColor = RGB(in2, in1, in2)
Next e
txtBox.SetFocus
End Sub

Private Sub cmdCutSpaces_Click()
Dim S$: S = RTrim(txtBox): txtBox = S
txtBox.SetFocus
End Sub

Private Sub cmdDouble_Click()
T = txtBox: L = Len(T): Dim C, P
For I = 1 To L
C = Mid(T, I, 1): P = P + C + C
Next I
txtBox = P: P = ""
txtBox.SetFocus
End Sub

Private Sub cmdDoubleRev_Click()
Dim X() As String, Y() As String, U#
X = Split(txtBox): U = UBound(X): ReDim Y(U)
For j = 0 To U: Y(U - j) = X(j): Next j
txtBox = Join(Y): SR$ = StrReverse(txtBox): txtBox = SR$
txtBox.SetFocus
End Sub

Private Sub cmdElite_Click()
Elite txtBox 'see module
txtBox.SetFocus
End Sub

Private Sub cmdExchange_Click()
On Error Resume Next
Dim N, B$, A$
N = InputBox("The position of the character to replace?", "Enter Position:")
N = CInt(N) 'convert to integer (if user was a retard...)
B = Mid(txtBox, 1, N - 1)
A = Mid(txtBox, N + 1, Len(txtBox) - N)
txtBox = B & A
txtBox.SetFocus
End Sub

Private Sub cmdHacker_Click()
T = txtBox: L = Len(T): Dim P: Dim C$: T = UCase(T)
For I = 1 To L
C = Mid(T, I, 1)
If C = "A" Or C = "E" Or C = "I" Or _
C = "O" Or C = "U" Then C = LCase(C)
P = P + C
Next I
txtBox = P: P = ""
txtBox.SetFocus
End Sub

Private Sub cmdIns_Click()
C = InputBox("Please enter what you'd like to insert:", "Insert Special")
T = txtBox: L = Len(T): Dim P
For I = 1 To L
S = Mid(T, I, 1): S = S + C: P = P + S
Next I
txtBox = P: P = ""
txtBox.SetFocus
End Sub

Private Sub cmdJumble_Click()
T = txtBox: L = Len(T): Dim C$, P$, B$, A$, r#
For I = 1 To L
r = Int(Rnd * L): r = r + 1:
C = Mid(T, r, 1): P = P + C
B = Mid(T, 1, r - 1): A = Mid(T, r + 1, L - r)
T = B & A: L = Len(T)
Next I
txtBox = P: P = ""
txtBox.SetFocus
End Sub

Private Sub cmdLamer_Click()
T = txtBox: L = Len(T): Dim P: Dim C$: T = LCase(T)
For I = 1 To L
C = Mid(T, I, 1)
If C = "a" Or C = "e" Or C = "i" Or _
C = "o" Or C = "u" Then C = UCase(C)
P = P + C
Next I
txtBox = P: P = ""
txtBox.SetFocus
End Sub

Private Sub cmdLoad_Click()
txtBox = Saved$
End Sub

Private Sub cmdLTrim_Click()
Dim S$: S = LTrim(txtBox): txtBox = S
txtBox.SetFocus
End Sub

Private Sub cmdReplace_Click()
Dim RWh, RWi, Final As String
RWh = InputBox("What letter or word to replace?", "Replace What?")
RWi = InputBox("What to replace it with?", "Replace With?")
If Len(RWi) = 0 Then RWi = ""
Final = Replace(txtBox, RWh, RWi)
txtBox = Final
txtBox.SetFocus
End Sub

Private Sub cmdRepOne_Click()
On Error Resume Next
Dim N, W$, B$, A$
N = InputBox("The position of the character to replace?", "Enter Position:")
N = CInt(N) 'convert to integer (if user was a retard...)
W = InputBox("What to replace character " & N & " with?", "Replace With?")
B = Mid(txtBox, 1, N - 1)
A = Mid(txtBox, N + 1, Len(txtBox) - N)
txtBox = B & W & A
txtBox.SetFocus
End Sub

Private Sub cmdRev_Click()
Dim S$: S = StrReverse(txtBox): txtBox = S
txtBox.SetFocus
End Sub

Private Sub cmdRevSen_Click()
Dim X() As String, Y() As String, U#
X = Split(txtBox): U = UBound(X): ReDim Y(U)
For j = 0 To U: Y(U - j) = X(j): Next j
txtBox = Join(Y)
txtBox.SetFocus
End Sub

Private Sub cmdSave_Click()
Saved$ = txtBox
End Sub

Private Sub cmdSpace_Click()
T = txtBox: L = Len(T): Dim P
For I = 1 To L
C = Mid(T, I, 1): C = C + " ": P = P + C
Next I
txtBox = P: P = ""
txtBox.SetFocus
End Sub

Private Sub cmdStrike_Click()
If txtBox.Font.Strikethrough = True Then txtBox.Font.Strikethrough = False Else txtBox.Font.Strikethrough = True
txtBox.SetFocus
End Sub

Private Sub cmdSwap_Click()
Dim O$, T$, S$
MsgBox "Enter a string, press OK, then enter another.  It will swap them.", vbInformation + vbOKOnly, "Swap"
O = InputBox("?", "Enter the first Swapee")
T = InputBox("?", "Enter the second Swapee")
S = Replace(txtBox, O, "&^&$$%^$#%!@#$")
txtBox = S
S = Replace(txtBox, T, O)
txtBox = S
S = Replace(txtBox, "&^&$$%^$#%!@#$", T)
txtBox = S
txtBox.SetFocus
End Sub

Private Sub cmdTypy_Click()
For Each object In frmMain: object.Enabled = False: Next
With txtBox
.Enabled = True: .Locked = True
Dim Str$: Str = .text: .text = ""
End With
For I = 1 To Len(Str)
Pause 0.1: txtBox = txtBox + Mid(Str, I, 1)
Next I
For Each object In frmMain: object.Enabled = True: Next
txtBox.SetFocus
txtBox.Locked = False
End Sub

Private Sub cmdULine_Click()
If txtBox.Font.Underline = True Then txtBox.Font.Underline = False Else txtBox.Font.Underline = True
txtBox.SetFocus
End Sub

Private Sub Form_Load()
Randomize
End Sub

Private Sub SCR_Click()
Dim F: L = Len(SCR.Caption)
For I = 1 To L
T = SCR.Caption: L = Len(SCR.Caption)
N = Int(Rnd * L): If N = 0 Or N = L - 1 Then N = N + 1
C = Mid(T, N, 1): F = F + C
j = Replace(SCR.Caption, C, ""): SCR.Caption = j
Next I
SCR.Caption = F: F = ""
End Sub
