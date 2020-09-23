VERSION 5.00
Begin VB.Form Form5 
   Caption         =   "MorphListBox - .RightToLeft Demonstration"
   ClientHeight    =   4305
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6120
   LinkTopic       =   "Form5"
   ScaleHeight     =   4305
   ScaleWidth      =   6120
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Height          =   975
      Left            =   3240
      TabIndex        =   2
      Top             =   2400
      Width           =   2775
      Begin VB.CheckBox Check4 
         Caption         =   "Large Size"
         Height          =   255
         Left            =   1560
         TabIndex        =   4
         Top             =   360
         Width           =   1095
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Display Icons"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Value           =   1  'Checked
         Width           =   1335
      End
   End
   Begin VB.CheckBox Check1 
      Caption         =   ".RightToLeft"
      Height          =   375
      Left            =   3960
      TabIndex        =   1
      Top             =   480
      Value           =   1  'Checked
      Width           =   1215
   End
   Begin prjMorphListBox.MorphListBox MorphListBox1 
      Height          =   3615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   6376
      CheckStyle      =   0
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
      MultiSelect     =   1
      RightToLeft     =   -1  'True
      SelColor1       =   6316128
      SelColor2       =   14737632
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Check1_Click()
   If Check1.Value = vbChecked Then
      MorphListBox1.RightToLeft = True
   Else
      MorphListBox1.RightToLeft = False
   End If
End Sub

Private Sub Check2_Click()
   'If Check2.Value = vbChecked Then
   '   MorphListBox1.Style = CheckBox
   '   Option1(0).Enabled = True
   '   Option1(1).Enabled = True
   '   Option1(2).Enabled = True
   'Else
   '   MorphListBox1.Style = Standard
   '   Option1(0).Enabled = False
   '   Option1(1).Enabled = False
   '   Option1(2).Enabled = False
   'End If
End Sub

Private Sub Check3_Click()
   If Check3.Value = vbChecked Then
      Check4.Enabled = True
      MorphListBox1.ShowItemImages = True
   Else
      Check4.Enabled = False
      MorphListBox1.ShowItemImages = False
   End If
End Sub

Private Sub Check4_Click()
   If Check4.Value = vbChecked Then
      MorphListBox1.ItemImageSize = 40
   Else
      MorphListBox1.ItemImageSize = 0 ' same size as font height
   End If
End Sub

Private Sub Form_Load()

   Dim i As Long, j As Long

   MorphListBox1.ShowItemImages = True

   MorphListBox1.AddImage App.Path & "\beavis.ico"      ' image index 0
   MorphListBox1.AddImage App.Path & "\butthead.ico"    ' image index 1
   MorphListBox1.AddImage App.Path & "\stimpy.ico"      ' image index 2
   MorphListBox1.AddImage App.Path & "\igor.ico"        ' image index 3
   MorphListBox1.AddImage App.Path & "\alien.ico"       ' image index 4
   MorphListBox1.AddImage App.Path & "\ren.ico"         ' image index 5

'  ALWAYS use .RedrawFlag when adding more than a few items.
   MorphListBox1.RedrawFlag = False
   j = 0
   For i = 1 To 100
      MorphListBox1.AddItem "0123456789"
      MorphListBox1.ImageIndex(MorphListBox1.NewIndex) = j
      j = j + 1
      If j > 5 Then j = 0
   Next i
   MorphListBox1.RedrawFlag = True

End Sub

Private Sub Option1_Click(Index As Integer)
   Select Case Index
      Case 0
         MorphListBox1.CheckStyle = Arrow
      Case 1
         MorphListBox1.CheckStyle = Tick
      Case 2
         MorphListBox1.CheckStyle = X
   End Select
End Sub
