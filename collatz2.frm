VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "collatz"
   ClientHeight    =   2595
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3570
   LinkTopic       =   "Form1"
   ScaleHeight     =   2595
   ScaleWidth      =   3570
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   2010
      Left            =   2160
      TabIndex        =   5
      Top             =   480
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   960
      TabIndex        =   3
      Top             =   480
      Width           =   1095
   End
   Begin VB.Timer Timer1 
      Left            =   720
      Top             =   1320
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   480
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "HESAPLA"
      Height          =   1575
      Left            =   -120
      TabIndex        =   0
      Top             =   840
      Width           =   2175
   End
   Begin VB.Label Label3 
      Caption         =   "sonuçlar"
      Height          =   255
      Left            =   2160
      TabIndex        =   6
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "'e kadar dene"
      Height          =   375
      Left            =   960
      TabIndex        =   4
      Top             =   240
      Width           =   975
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
Dim i, j, ilk, son, sayac As Double
Private Sub Command1_Click()
List1.Clear
i = Text1.Text
sayac = 0
ilk = Text1.Text
son = Text2.Text

For sayac = ilk To son
i = sayac

Hesap:
j = j + 1
If i Mod 2 = 0 Then
i = i / 2
Else
i = (3 * i) + 1
End If

If i = 1 Then
'MsgBox (j & " Adýmda Bitti")
List1.AddItem j
Call SaveLoadListbox(List1, "collatz2.txt", "save")
Else
GoTo Hesap
End If

i = 0
j = 1
Next

MsgBox ("bitti")

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







