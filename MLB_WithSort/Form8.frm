VERSION 5.00
Begin VB.Form Form8 
   Caption         =   "MorphListBox - Sorting"
   ClientHeight    =   6075
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   11340
   LinkTopic       =   "Form8"
   ScaleHeight     =   6075
   ScaleWidth      =   11340
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "Sort"
      Height          =   375
      Left            =   4560
      TabIndex        =   12
      Top             =   5280
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   6000
      TabIndex        =   10
      Text            =   "1000"
      Top             =   870
      Width           =   735
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Populate List"
      Height          =   375
      Left            =   4560
      TabIndex        =   9
      Top             =   840
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Sort"
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   5280
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1560
      TabIndex        =   2
      Text            =   "1000"
      Top             =   870
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Populate List"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   1335
   End
   Begin prjMorphListBox.MorphListBox mlb 
      Height          =   3615
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   6376
      BeginProperty ListFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Sorted          =   0   'False
   End
   Begin prjMorphListBox.MorphListBox mlb2 
      Height          =   3615
      Left            =   4560
      TabIndex        =   6
      Top             =   1320
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   6376
      BeginProperty ListFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SortAsNumeric   =   -1  'True
      Sorted          =   0   'False
   End
   Begin VB.Label Label7 
      Caption         =   $"Form8.frx":0000
      Height          =   3015
      Left            =   8040
      TabIndex        =   14
      Top             =   1560
      Width           =   3015
   End
   Begin VB.Label Label6 
      Caption         =   "Label2"
      Height          =   375
      Left            =   6120
      TabIndex        =   13
      Top             =   5280
      Width           =   1575
   End
   Begin VB.Label Label5 
      Caption         =   "Label1"
      Height          =   255
      Left            =   6840
      TabIndex        =   11
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "Sorting as Numerics"
      Height          =   255
      Left            =   4680
      TabIndex        =   8
      Top             =   360
      Width           =   2895
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Sorting as Strings"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   360
      Width           =   2895
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   375
      Left            =   1680
      TabIndex        =   5
      Top             =   5280
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Left            =   2400
      TabIndex        =   4
      Top             =   840
      Width           =   1335
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
   Dim Stopwatch As New CStopWatch
   Private testarray() As String

Private Sub Command1_Click()
   Dim i As Long, X As Long
   mlb.Clear
   GenerateRandomList
End Sub

Private Sub Command2_Click()
Stopwatch.Reset
mlb.Sort
Label2.Caption = Stopwatch.Elapsed & " ms"
End Sub

Private Sub Command3_Click()
   Dim Stopwatch As New CStopWatch
   
   Dim i As Long, X As Long
   mlb2.Clear
   mlb2.RedrawFlag = False
   X = Val(Text2.Text)
   Stopwatch.Reset
   For i = 1 To X
      mlb2.AddItem CStr(Int(Rnd(1) * 1000000) + 1)
   Next i
   mlb2.RedrawFlag = True
   Label5.Caption = CStr(Stopwatch.Elapsed) & " ms"
End Sub

Private Sub Command4_Click()
Stopwatch.Reset
mlb2.Sort
Label6.Caption = Stopwatch.Elapsed & " ms"

End Sub

Private Sub Form_Load()
   Randomize
End Sub

Private Sub GenerateRandomList()

   Dim Stopwatch As New CStopWatch
   Dim Max As Long, i As Long, j As Long, k As Long, s As String, letter As Integer

   ReDim testarray(1 To Val(Text1.Text))
   Max = Val(Text1.Text)

'  generate a list of 'max' random strings
   For i = 1 To Max
      k = Int(Rnd * 18) + 1
      s = ""
      For j = 1 To k '30
         letter = Int((90 - 65 + 1) * Rnd + 65)
         s = s & Chr(letter)
      Next j
      testarray(i) = s
   Next i

   Stopwatch.Reset

   mlb.RedrawFlag = False
   For j = 1 To Max
      mlb.AddItem testarray(j)
   Next j
   mlb.RedrawFlag = True


   Label1.Caption = CStr(Stopwatch.Elapsed) & " ms"

End Sub

