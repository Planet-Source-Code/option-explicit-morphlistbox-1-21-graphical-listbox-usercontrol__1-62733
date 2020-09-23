VERSION 5.00
Begin VB.Form Form3 
   AutoRedraw      =   -1  'True
   Caption         =   "MorphListBox - Matthew R. Usner - .SortAsNumeric"
   ClientHeight    =   3660
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6015
   LinkTopic       =   "Form3"
   ScaleHeight     =   3660
   ScaleWidth      =   6015
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Exit"
      Height          =   375
      Left            =   2400
      TabIndex        =   3
      Top             =   3240
      Width           =   1215
   End
   Begin prjMorphListBox.MorphListBox MorphListBox2 
      Height          =   2175
      Left            =   3120
      TabIndex        =   1
      Top             =   120
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   3836
      BackColor2      =   16744703
      BackColor1      =   8388736
      BorderColor     =   4194368
      BeginProperty ListFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SelColor1       =   8388736
      SelColor2       =   8388736
      SelTextColor    =   16761087
      TrackBarColor1  =   8388736
      TrackBarColor2  =   16761087
      ButtonColor1    =   4194368
      ButtonColor2    =   8388736
      ThumbColor1     =   4194368
      ThumbColor2     =   8388736
      ThumbBorderColor=   16711935
      ArrowUpColor    =   16711935
      ArrowDownColor  =   8388736
      Theme           =   6
      CheckboxArrowColor=   4194368
      CheckBoxColor   =   4194368
      FocusRectColor  =   16761087
      TrackClickColor1=   4194368
      TrackClickColor2=   16744703
      SortAsNumeric   =   -1  'True
   End
   Begin prjMorphListBox.MorphListBox MorphListBox1 
      Height          =   2175
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   3836
      BackColor2      =   16777088
      BackColor1      =   8421376
      BorderColor     =   4210688
      BeginProperty ListFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SelColor1       =   8421376
      SelColor2       =   8421376
      SelTextColor    =   16777152
      TrackBarColor1  =   8421376
      TrackBarColor2  =   16777152
      ButtonColor1    =   4210688
      ButtonColor2    =   8421376
      ThumbColor1     =   4210688
      ThumbColor2     =   8421376
      ThumbBorderColor=   16776960
      ArrowUpColor    =   16776960
      ArrowDownColor  =   4210688
      Theme           =   1
      CheckboxArrowColor=   4210688
      CheckBoxColor   =   4210688
      FocusRectColor  =   16777152
      TrackClickColor1=   4210688
      TrackClickColor2=   16777088
   End
   Begin VB.Label Label1 
      Caption         =   $"Form3.frx":0000
      Height          =   1095
      Left            =   120
      TabIndex        =   2
      Top             =   2400
      Width           =   5775
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()

   Dim TestArray(1 To 100) As String

   For i = 1 To 100
      k = Int(Rnd(1) * 50) + 1
      TestArray(i) = CStr(k)
   Next
   
   MorphListBox1.RedrawFlag = False
   MorphListBox2.RedrawFlag = False

   For i = 1 To 100
      MorphListBox1.AddItem TestArray(i)
      MorphListBox2.AddItem TestArray(i)
   Next i

   MorphListBox1.RedrawFlag = True
   MorphListBox2.RedrawFlag = True

End Sub
