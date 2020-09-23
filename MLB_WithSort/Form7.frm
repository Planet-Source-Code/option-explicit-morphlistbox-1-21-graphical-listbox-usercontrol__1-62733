VERSION 5.00
Begin VB.Form Form7 
   Caption         =   "MorphListBox - Themes"
   ClientHeight    =   4170
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6045
   LinkTopic       =   "Form7"
   ScaleHeight     =   4170
   ScaleWidth      =   6045
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Populate List"
      Height          =   375
      Left            =   3480
      TabIndex        =   11
      Top             =   3600
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   4920
      TabIndex        =   10
      Text            =   "1000"
      Top             =   3600
      Width           =   735
   End
   Begin VB.Frame Frame1 
      Height          =   3135
      Left            =   3480
      TabIndex        =   1
      Top             =   360
      Width           =   2175
      Begin VB.OptionButton Option1 
         Caption         =   "Blue Moon"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Cyan Eyed"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   600
         Width           =   1215
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Golden Goose"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   960
         Width           =   1455
      End
      Begin VB.OptionButton Option4 
         Caption         =   "Green With Envy"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   1320
         Width           =   1695
      End
      Begin VB.OptionButton Option5 
         Caption         =   "Gunmetal Grey"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   1680
         Width           =   1455
      End
      Begin VB.OptionButton Option6 
         Caption         =   "Penny Wise"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   2040
         Width           =   1215
      End
      Begin VB.OptionButton Option7 
         Caption         =   "Purple People Eater"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   2400
         Width           =   1815
      End
      Begin VB.OptionButton Option8 
         Caption         =   "Red Rum"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   2760
         Width           =   1215
      End
   End
   Begin prjMorphListBox.MorphListBox mlb 
      Height          =   3495
      Left            =   120
      TabIndex        =   0
      Top             =   240
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
      Theme           =   3
      ThumbBorderColor=   16761024
      ThumbColor1     =   4194304
      ThumbColor2     =   16744576
      TrackBarColor1  =   8388608
      TrackBarColor2  =   16761024
      TrackClickColor1=   4194304
      TrackClickColor2=   16744576
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
   Dim i As Long, X As Long
   mlb.Clear
   mlb.RedrawFlag = False
   X = Val(Text1.Text)
   For i = 1 To X
      mlb.AddItem CStr(Int(Rnd(1) * 1000000) + 1)
   Next i
   mlb.RedrawFlag = True
End Sub

Private Sub Form_Load()
   Randomize
End Sub

Private Sub Option1_Click()
If Option1.Value Then mlb.Theme = [Blue Moon]
End Sub

Private Sub Option2_Click()
If Option2.Value Then mlb.Theme = [Cyan Eyed]
End Sub

Private Sub Option3_Click()
If Option3.Value Then mlb.Theme = [Golden Goose]
End Sub

Private Sub Option4_Click()
If Option4.Value Then mlb.Theme = [Green With Envy]
End Sub

Private Sub Option5_Click()
If Option5.Value Then mlb.Theme = [Gunmetal Grey]
End Sub

Private Sub Option6_Click()
If Option6.Value Then mlb.Theme = [Penny Wise]
End Sub

Private Sub Option7_Click()
If Option7.Value Then mlb.Theme = [Purple People Eater]
End Sub

Private Sub Option8_Click()
If Option8.Value Then mlb.Theme = [Red Rum]
End Sub

