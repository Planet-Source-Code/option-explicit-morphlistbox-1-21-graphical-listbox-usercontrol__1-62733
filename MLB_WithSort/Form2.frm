VERSION 5.00
Begin VB.Form Form2 
   AutoRedraw      =   -1  'True
   Caption         =   "MorphListBox - Matthew R. Usner - Drag and Drop"
   ClientHeight    =   3885
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6015
   LinkTopic       =   "Form2"
   ScaleHeight     =   3885
   ScaleWidth      =   6015
   StartUpPosition =   3  'Windows Default
   Begin prjMorphListBox.MorphListBox MorphListBox1 
      Height          =   2295
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   4048
      BackColor2      =   12648384
      BackColor1      =   32768
      BorderColor     =   16384
      ListIndex       =   0
      BeginProperty ListFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SelColor1       =   32768
      SelColor2       =   32768
      SelTextColor    =   12648384
      MultiSelect     =   1
      TrackBarColor1  =   32768
      TrackBarColor2  =   12648384
      ButtonColor1    =   16384
      ButtonColor2    =   8454016
      ThumbColor1     =   16384
      ThumbColor2     =   65280
      ThumbBorderColor=   12648384
      ArrowUpColor    =   12648384
      ArrowDownColor  =   16384
      Theme           =   5
      CheckboxArrowColor=   16384
      CheckBoxColor   =   16384
      FocusRectColor  =   12648384
      DragEnabled     =   -1  'True
      TrackClickColor1=   16384
      TrackClickColor2=   8454016
   End
   Begin VB.CheckBox chkRemove 
      Caption         =   "Remove From Source ListBox"
      Height          =   195
      Left            =   240
      TabIndex        =   3
      Top             =   3600
      Value           =   1  'Checked
      Width           =   3375
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   4080
      TabIndex        =   0
      Top             =   3360
      Width           =   1695
   End
   Begin prjMorphListBox.MorphListBox MorphListBox2 
      Height          =   2295
      Left            =   3120
      TabIndex        =   1
      Top             =   120
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   4048
      BackColor2      =   8438015
      BackColor1      =   16512
      BorderColor     =   4210816
      BeginProperty ListFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SelColor1       =   16512
      SelColor2       =   16512
      SelTextColor    =   12640511
      TrackBarColor1  =   16512
      TrackBarColor2  =   12640511
      ButtonColor1    =   4210816
      ButtonColor2    =   33023
      ThumbColor1     =   4210816
      ThumbColor2     =   16576
      ThumbBorderColor=   33023
      ArrowUpColor    =   12640511
      ArrowDownColor  =   16512
      Theme           =   8
      FocusRectColor  =   12640511
      TrackClickColor1=   4210816
      TrackClickColor2=   8438015
   End
   Begin VB.Label Label1 
      Caption         =   $"Form2.frx":0000
      Height          =   975
      Left            =   120
      TabIndex        =   2
      Top             =   2520
      Width           =   5775
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdExit_Click()
   Unload Me
End Sub

Private Sub Form_Load()

   Dim i As Long

   MorphListBox1.RedrawFlag = False ' ".RedrawFlag is a loop's best friend".
   For i = 1 To 50
      MorphListBox1.AddItem String(20, Chr(64 + i))
   Next
   MorphListBox1.RedrawFlag = True

End Sub

Private Sub MorphListBox1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

   If MorphListBox1.DragEnabled Then
     MorphListBox1.DragIcon = LoadPicture(App.Path & "\" & "drag1pg.ico")
     MorphListBox1.Drag vbBeginDrag
   End If

End Sub

Private Sub MorphListBox2_DragDrop(Source As Control, X As Single, Y As Single)

'  the source listbox in this example is set to MultiSelect Simple.
   Dim i As Long, lCount As Long

'  notice I use the RedrawFlag property for both the source and target
'  listboxes.  Since removing and adding items necessarily involves
'  redrawing the control, suspending all redrawing until after all drag
'  and drop operations are complete vastly increases speed.  The moral?
'  use .RedrawFlag whenever there's a possibility that there will be many
'  additions or removals from listboxes.  Repeat after me:
'  ".RedrawFlag is a loop's best friend."
   Source.RedrawFlag = False
   MorphListBox2.RedrawFlag = False

   i = 0
   lCount = Source.ListCount
   While i <= lCount - 1
      If Source.Selected(i) Then
         MorphListBox2.AddItem Source.List(i)
         If chkRemove.Value = vbChecked Then
            Source.RemoveItem i
            lCount = lCount - 1
            i = i - 1
         End If
      End If
      i = i + 1
   Wend

   Source.RedrawFlag = True
   MorphListBox2.RedrawFlag = True

End Sub
