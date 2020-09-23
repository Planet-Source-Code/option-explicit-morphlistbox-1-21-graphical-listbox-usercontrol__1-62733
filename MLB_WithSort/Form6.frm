VERSION 5.00
Begin VB.Form Form6 
   Caption         =   "MorphListBox - Speed Comparison; ClearSelected Demo"
   ClientHeight    =   6135
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7455
   LinkTopic       =   "Form6"
   ScaleHeight     =   6135
   ScaleWidth      =   7455
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   ".ClearOrSelect Method Test"
      Height          =   1095
      Left            =   240
      TabIndex        =   8
      Top             =   4920
      Width           =   7095
      Begin VB.CheckBox chkSelect 
         Caption         =   "Select"
         Height          =   255
         Left            =   3000
         TabIndex        =   16
         ToolTipText     =   "Check this box to select all items in supplied range."
         Top             =   480
         Width           =   855
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Clear/Select"
         Height          =   375
         Left            =   3960
         TabIndex        =   11
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox txtStart 
         Height          =   285
         Left            =   720
         TabIndex        =   10
         Text            =   "-1"
         Top             =   450
         Width           =   735
      End
      Begin VB.TextBox txtEnd 
         Height          =   285
         Left            =   2040
         TabIndex        =   9
         Text            =   "-1"
         Top             =   450
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Start"
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   480
         Width           =   495
      End
      Begin VB.Label Label4 
         Caption         =   "End"
         Height          =   255
         Left            =   1680
         TabIndex        =   14
         Top             =   480
         Width           =   375
      End
      Begin VB.Label Label5 
         Caption         =   "Num Selected"
         Height          =   255
         Left            =   5760
         TabIndex        =   13
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label lblNumSelected 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   5640
         TabIndex        =   12
         Top             =   660
         Width           =   1215
      End
   End
   Begin VB.ListBox List1 
      Height          =   3495
      IntegralHeight  =   0   'False
      Left            =   4320
      Sorted          =   -1  'True
      TabIndex        =   3
      Top             =   120
      Width           =   3015
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   3840
      TabIndex        =   2
      Text            =   "1000"
      Top             =   4320
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Populate Lists"
      Height          =   375
      Left            =   2280
      TabIndex        =   1
      Top             =   4320
      Width           =   1335
   End
   Begin prjMorphListBox.MorphListBox mlb 
      Height          =   3495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   6165
      ArrowUpColor    =   16761024
      ArrowDownColor  =   4194304
      BackColor2      =   16761024
      BackColor1      =   12582912
      BorderColor1    =   4194304
      BorderColor2    =   16744576
      BorderWidth     =   16
      ButtonColor1    =   4194304
      ButtonColor2    =   16744576
      CheckboxArrowColor=   4194304
      CheckBoxColor   =   4194304
      FocusBorderColor1=   4194304
      FocusBorderColor2=   16711680
      FocusRectColor  =   16761024
      BeginProperty ListFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SelColor1       =   8388608
      SelColor2       =   8388608
      SelTextColor    =   16761024
      Style           =   1
      Theme           =   3
      ThumbBorderColor=   16761024
      ThumbColor1     =   4194304
      ThumbColor2     =   16744576
      TrackBarColor1  =   8388608
      TrackBarColor2  =   16761024
      TrackClickColor1=   4194304
      TrackClickColor2=   16744576
   End
   Begin VB.Label lblTimeVB 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   5400
      TabIndex        =   7
      Top             =   3840
      Width           =   1215
   End
   Begin VB.Label lblTimeMLB 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1200
      TabIndex        =   6
      Top             =   3840
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Time (ms):"
      Height          =   255
      Left            =   4320
      TabIndex        =   5
      Top             =   3840
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Time (ms):"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   3840
      Width           =   855
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
   Dim Stopwatch As New CStopWatch

Private Sub Command1_Click()

   Dim i As Long, X As Long
   mlb.Clear
   mlb.RedrawFlag = False
   X = Val(Text1.Text)
   Stopwatch.Reset
   For i = 1 To X
      mlb.AddItem CStr(Int(Rnd(1) * 1000000) + 1)
   Next i
   lblTimeMLB.Caption = Stopwatch.Elapsed

   List1.Clear
   Stopwatch.Reset
   For i = 1 To X
      List1.AddItem CStr(Int(Rnd(1) * 1000000) + 1)
   Next i
   lblTimeVB.Caption = Stopwatch.Elapsed

   mlb.RedrawFlag = True
End Sub

Private Sub Command2_Click()
   If chkSelect.Value = vbChecked Then
      mlb.ClearOrSelect Val(txtStart.Text), Val(txtEnd.Text), True
   Else
      mlb.ClearOrSelect Val(txtStart.Text), Val(txtEnd.Text), False
   End If
   lblNumSelected.Caption = CStr(mlb.SelCount)
End Sub

Private Sub Form_Load()
   Randomize
End Sub

Private Sub mlb_KeyUp(KeyCode As Integer, Shift As Integer)
   lblNumSelected.Caption = CStr(mlb.SelCount)
End Sub

Private Sub mlb_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   lblNumSelected.Caption = CStr(mlb.SelCount)
End Sub

