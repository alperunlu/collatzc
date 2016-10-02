VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "collatz"
   ClientHeight    =   2460
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   2460
   LinkTopic       =   "Form1"
   ScaleHeight     =   2460
   ScaleWidth      =   2460
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   2010
      Left            =   1080
      TabIndex        =   3
      Top             =   480
      Width           =   1335
   End
   Begin VB.Timer Timer1 
      Left            =   1320
      Top             =   1560
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   480
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "SONUÇ"
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "sonuçlar"
      Height          =   255
      Left            =   1080
      TabIndex        =   4
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "'den baþla"
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   240
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim i, j As Double
Private Sub Command1_Click()
List1.Clear
List1.AddItem Text1.Text
i = Text1.Text


Hesap:
j = j + 1
If i Mod 2 = 0 Then
i = i / 2
List1.AddItem i
Else
i = (3 * i) + 1
List1.AddItem i
End If


If i = 1 Then

MsgBox (j & " Adýmda Bitti")
Call SaveLoadListbox(List1, "collatz.txt", "save")
i = 0
j = 1
Else
GoTo Hesap
End If
End Sub
Private Sub Form_Load()
i = 0
j = 1
End Sub
Private Sub SaveLoadListbox(plstLB As ListBox, pstrFileName As String, _
pstrSaveOrLoad As String)

Dim strListItems As String
Dim i As Long

Select Case pstrSaveOrLoad
   Case "save"
    Open pstrFileName For Output As #1
    For i = 0 To plstLB.ListCount - 1
        plstLB.Selected(i) = True
        Print #1, plstLB.List(plstLB.ListIndex)
    Next
    Close #1

   Case "load"
   plstLB.Clear
    Open pstrFileName For Input As #1
    While Not EOF(1)
      Line Input #1, strListItems
      plstLB.AddItem strListItems
    Wend
    Close #1
End Select

End Sub





