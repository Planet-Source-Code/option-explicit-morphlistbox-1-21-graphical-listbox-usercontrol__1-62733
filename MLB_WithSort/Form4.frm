VERSION 5.00
Begin VB.Form Form4 
   Caption         =   "MorphListBox - Icon Display Demonstration - Matthew R. Usner"
   ClientHeight    =   4575
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6615
   LinkTopic       =   "Form4"
   ScaleHeight     =   4575
   ScaleWidth      =   6615
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check4 
      Caption         =   "Enabled"
      Height          =   375
      Left            =   4680
      TabIndex        =   4
      Top             =   720
      Value           =   1  'Checked
      Width           =   1815
   End
   Begin VB.CheckBox Check3 
      Caption         =   "Large Icons"
      Height          =   255
      Left            =   4680
      TabIndex        =   3
      Top             =   480
      Value           =   1  'Checked
      Width           =   1575
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Show ListItem Icons"
      Height          =   255
      Left            =   4680
      TabIndex        =   2
      Top             =   120
      Value           =   1  'Checked
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4920
      TabIndex        =   1
      Top             =   3960
      Width           =   1335
   End
   Begin prjMorphListBox.MorphListBox MorphListBox1 
      Height          =   4335
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   7646
      ArrowUpColor    =   12648384
      ArrowDownColor  =   16384
      BackColor2      =   0
      BackColor1      =   16761024
      BorderColor1    =   4194304
      BorderColor2    =   12582912
      BorderWidth     =   16
      ButtonColor1    =   16384
      ButtonColor2    =   8454016
      CheckboxArrowColor=   12648447
      CheckBoxColor   =   65535
      CircularGradient=   -1  'True
      FocusBorderColor1=   4194304
      FocusBorderColor2=   8388608
      FocusRectColor  =   12648384
      ItemImageSize   =   40
      BeginProperty ListFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SelColor1       =   32768
      SelColor2       =   32768
      SelTextColor    =   12648384
      TextColor       =   12648384
      Theme           =   5
      ThumbBorderColor=   12648384
      ThumbColor1     =   16384
      ThumbColor2     =   65280
      TrackBarColor1  =   32768
      TrackBarColor2  =   12648384
      TrackClickColor1=   16384
      TrackClickColor2=   8454016
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()
   If Check1.Value = vbChecked Then
      MorphListBox1.ShowItemImages = True
   Else
      MorphListBox1.ShowItemImages = False
   End If
End Sub

Private Sub Check3_Click()
   If Check3.Value = vbChecked Then
      MorphListBox1.ItemImageSize = 40
   Else
      MorphListBox1.ItemImageSize = 0 ' same size as font height
   End If

End Sub

Private Sub Check4_Click()
   If Check4.Value = vbChecked Then
      MorphListBox1.Enabled = True
   Else
      MorphListBox1.Enabled = False
   End If
End Sub

Private Sub Command1_Click()
   Unload Me
End Sub

Private Sub Form_Load()

   MorphListBox1.ShowItemImages = True

   MorphListBox1.AddImage App.Path & "\beavis.ico"      ' image index 0
   MorphListBox1.AddImage App.Path & "\butthead.ico"    ' image index 1
   MorphListBox1.AddImage App.Path & "\stimpy.ico"      ' image index 2
   MorphListBox1.AddImage App.Path & "\igor.ico"        ' image index 3
   MorphListBox1.AddImage App.Path & "\alien.ico"       ' image index 4
   MorphListBox1.AddImage App.Path & "\ren.ico"         ' image index 5

'  no offense guys, it's just a demo. :)

'  even though I'm just adding a few items here, it's a good idea to get
'  into the habit of using .RedrawFlag when adding or removing listitems.
   MorphListBox1.RedrawFlag = False

   MorphListBox1.AddItem "LaVolpe"
   MorphListBox1.ImageIndex(MorphListBox1.NewIndex) = 3 ' LaVolpe = Igor

   MorphListBox1.AddItem "Matthew R. Usner"
   MorphListBox1.ImageIndex(MorphListBox1.NewIndex) = 1 ' Matt = Butthead (ahem...)
   
   MorphListBox1.AddItem "Richard Mewett"
   MorphListBox1.ImageIndex(MorphListBox1.NewIndex) = 0 ' Richard = Beavis

   MorphListBox1.AddItem "Jim Jose"
   MorphListBox1.ImageIndex(MorphListBox1.NewIndex) = 2 ' Jim = Stimpy

   MorphListBox1.AddItem "Heriberto Mantilla Santamaria"
   MorphListBox1.ImageIndex(MorphListBox1.NewIndex) = 4 ' Heriberto = Alien

   MorphListBox1.AddItem "Carles P.V."
   MorphListBox1.ImageIndex(MorphListBox1.NewIndex) = 5 ' Carles = Ren

   MorphListBox1.RedrawFlag = True

End Sub
