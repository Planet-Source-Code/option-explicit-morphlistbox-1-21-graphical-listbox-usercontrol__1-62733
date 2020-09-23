VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "MorphListBox Demo - Matthew R. Usner"
   ClientHeight    =   8430
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8535
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8430
   ScaleWidth      =   8535
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkShowItemRect 
      Caption         =   "ListItem Focus Rectangle"
      Height          =   255
      Left            =   2640
      TabIndex        =   48
      Top             =   8040
      Value           =   1  'Checked
      Width           =   2175
   End
   Begin VB.Frame Frame8 
      Caption         =   ".TopIndex Property Usage"
      Height          =   615
      Left            =   5280
      TabIndex        =   46
      Top             =   1320
      Width           =   3135
      Begin VB.Label lblTopIndex 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   47
         Top             =   240
         Width           =   2895
      End
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Enabled"
      Height          =   255
      Left            =   120
      TabIndex        =   45
      Top             =   8040
      Value           =   1  'Checked
      Width           =   2175
   End
   Begin VB.Frame Frame7 
      Caption         =   ".DisplayFrom Method"
      Height          =   615
      Left            =   5280
      TabIndex        =   42
      Top             =   6720
      Width           =   3135
      Begin VB.CommandButton cmdDisplayFrom 
         Caption         =   "Do It"
         Height          =   255
         Left            =   2160
         TabIndex        =   44
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox txtDisplayFrom 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   43
         Text            =   "0"
         ToolTipText     =   "Enter a list item index value."
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Item Mouse is Over  (.MouseOverIndex)"
      Height          =   615
      Left            =   5280
      TabIndex        =   38
      Top             =   2040
      Width           =   3135
      Begin VB.Label lblMouseOverItem 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   39
         Top             =   240
         Width           =   2895
      End
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   495
      Left            =   5280
      TabIndex        =   37
      Top             =   7680
      Width           =   3135
   End
   Begin VB.Frame Frame5 
      Caption         =   ".RemoveItem Method"
      Height          =   615
      Left            =   5280
      TabIndex        =   34
      Top             =   4920
      Width           =   3135
      Begin VB.TextBox txtRemoveItem 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   36
         Text            =   "0"
         ToolTipText     =   "Enter index of item to remove."
         Top             =   240
         Width           =   1935
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Remove"
         Height          =   255
         Left            =   2160
         TabIndex        =   35
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   ".AddItem Method"
      Height          =   975
      Left            =   5280
      TabIndex        =   31
      Top             =   3840
      Width           =   3135
      Begin VB.TextBox txtAddIndex 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   720
         TabIndex        =   41
         Text            =   "-1"
         ToolTipText     =   "Type an optional index to insert item in."
         Top             =   600
         Width           =   975
      End
      Begin VB.CommandButton cmdAddItem 
         Caption         =   "Add"
         Height          =   255
         Left            =   2160
         TabIndex        =   33
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox txtAddItem 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   32
         ToolTipText     =   "Type a string to add to the list."
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label Label10 
         Caption         =   "Index:"
         Height          =   255
         Left            =   120
         TabIndex        =   40
         Top             =   600
         Width           =   495
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Load All ListBoxes"
      Height          =   975
      Left            =   5280
      TabIndex        =   26
      Top             =   2760
      Width           =   3135
      Begin VB.CommandButton cmdClear 
         Caption         =   "Clear"
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdLoad 
         Caption         =   "Load"
         Height          =   255
         Left            =   1080
         TabIndex        =   28
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox txtAmount 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2040
         TabIndex        =   27
         Text            =   "100"
         Top             =   240
         Width           =   975
      End
      Begin VB.Label lblUCLoadTime 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   600
         Width           =   2895
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   ".FindIndex Method"
      Height          =   975
      Left            =   5280
      TabIndex        =   20
      Top             =   5640
      Width           =   3135
      Begin VB.CheckBox chkCase 
         Caption         =   "Case Sensitive"
         Height          =   255
         Left            =   1560
         TabIndex        =   25
         ToolTipText     =   "Check if you wish the search to be case sensitive."
         Top             =   600
         Width           =   1455
      End
      Begin VB.CommandButton cmdFind 
         Caption         =   "Find"
         Height          =   255
         Left            =   2160
         TabIndex        =   22
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox txtFind 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   21
         ToolTipText     =   "Enter a string to search for in the listbox."
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label lblFind 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   720
         TabIndex        =   24
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label8 
         Caption         =   "Index:"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   600
         Width           =   495
      End
   End
   Begin VB.OptionButton optStats 
      Caption         =   "Property Stats"
      Height          =   255
      Index           =   4
      Left            =   3240
      TabIndex        =   14
      Top             =   7680
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      Caption         =   "ListBox Property Stats"
      Height          =   1095
      Left            =   5280
      TabIndex        =   11
      Top             =   120
      Width           =   3135
      Begin VB.Label lblCountOK 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2160
         TabIndex        =   19
         Top             =   480
         Width           =   735
      End
      Begin VB.Label lblListIndex 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1200
         TabIndex        =   18
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label Label9 
         Caption         =   "ListIndex"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   720
         Width           =   855
      End
      Begin VB.Label lblSelCount 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1200
         TabIndex        =   16
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label7 
         Caption         =   "SelCount"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   480
         Width           =   855
      End
      Begin VB.Label lblListCount 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1200
         TabIndex        =   13
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label5 
         Caption         =   "ListCount"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.OptionButton optStats 
      Caption         =   "Property Stats"
      Height          =   255
      Index           =   3
      Left            =   480
      TabIndex        =   10
      Top             =   7680
      Width           =   1695
   End
   Begin VB.OptionButton optStats 
      Caption         =   "Property Stats"
      Height          =   255
      Index           =   2
      Left            =   3000
      TabIndex        =   9
      Top             =   3720
      Width           =   1695
   End
   Begin VB.OptionButton optStats 
      Caption         =   "Property Stats"
      Height          =   255
      Index           =   1
      Left            =   480
      TabIndex        =   8
      Top             =   3720
      Value           =   -1  'True
      Width           =   1695
   End
   Begin prjMorphListBox.MorphListBox ucMorphListBox1 
      Height          =   3255
      Index           =   3
      Left            =   120
      TabIndex        =   2
      Top             =   4080
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   5741
      ArrowUpColor    =   16761024
      ArrowDownColor  =   4194304
      BackColor2      =   16761024
      BackColor1      =   12582912
      ButtonColor1    =   4194304
      ButtonColor2    =   16744576
      CheckboxArrowColor=   4194304
      CheckBoxColor   =   4194304
      DisPicture      =   "Form1.frx":0000
      FocusRectColor  =   16761024
      ListIndex       =   0
      BeginProperty ListFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MultiSelect     =   2
      Picture         =   "Form1.frx":BDF4
      SelColor1       =   8388608
      SelColor2       =   8388608
      SelTextColor    =   16761024
      TextColor       =   16776960
      Theme           =   3
      ThumbBorderColor=   16761024
      ThumbColor1     =   4194304
      ThumbColor2     =   16744576
      TrackBarColor1  =   8388608
      TrackBarColor2  =   16761024
      TrackClickColor1=   4194304
      TrackClickColor2=   16744576
   End
   Begin prjMorphListBox.MorphListBox ucMorphListBox1 
      Height          =   3255
      Index           =   2
      Left            =   2640
      TabIndex        =   1
      Top             =   120
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   5741
      ArrowUpColor    =   65535
      ArrowDownColor  =   32896
      BackColor2      =   8454143
      BackColor1      =   32896
      ButtonColor1    =   16448
      ButtonColor2    =   32896
      CheckboxArrowColor=   16448
      CheckBoxColor   =   16448
      CircularGradient=   -1  'True
      FocusRectColor  =   12648447
      BeginProperty ListFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SelColor1       =   32896
      SelColor2       =   32896
      SelTextColor    =   12648447
      Theme           =   7
      ThumbBorderColor=   65535
      ThumbColor1     =   16448
      ThumbColor2     =   32896
      TrackBarColor1  =   32896
      TrackBarColor2  =   12648447
      TrackClickColor1=   16448
      TrackClickColor2=   8454143
   End
   Begin prjMorphListBox.MorphListBox ucMorphListBox1 
      Height          =   3255
      Index           =   1
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   5741
      BorderColor2    =   14737632
      BorderWidth     =   7
      FocusBorderColor2=   8421504
      FocusRectColor  =   16777215
      BeginProperty ListFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SelColor1       =   4210752
      SelColor2       =   4210752
      SelTextColor    =   14737632
      Style           =   1
   End
   Begin prjMorphListBox.MorphListBox ucMorphListBox1 
      Height          =   3255
      Index           =   4
      Left            =   2640
      TabIndex        =   3
      Top             =   4080
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   5741
      ArrowUpColor    =   12632319
      ArrowDownColor  =   64
      BackColor2      =   12632319
      BackColor1      =   128
      BorderColor1    =   64
      BorderColor2    =   8421631
      BorderWidth     =   9
      ButtonColor1    =   64
      ButtonColor2    =   8421631
      CheckboxArrowColor=   64
      CheckBoxColor   =   64
      DisPicture      =   "Form1.frx":1A41E
      DisPictureMode  =   2
      FocusBorderColor1=   64
      FocusBorderColor2=   192
      FocusRectColor  =   12632319
      ListIndex       =   0
      BeginProperty ListFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MultiSelect     =   1
      Picture         =   "Form1.frx":1A770
      PictureMode     =   2
      SelColor1       =   128
      SelColor2       =   8421631
      SelTextColor    =   16777215
      TextColor       =   65535
      Theme           =   4
      ThumbBorderColor=   12632319
      ThumbColor1     =   64
      ThumbColor2     =   8421631
      TrackBarColor1  =   128
      TrackBarColor2  =   12632319
      TrackClickColor1=   64
      TrackClickColor2=   8421631
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "MultiSelect Extended, BitMap"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   7440
      Width           =   2415
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "MultiSelect Simple, Tiled Bitmap"
      Height          =   255
      Left            =   2760
      TabIndex        =   6
      Top             =   7440
      Width           =   2415
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "MultiSelect None, Golden Goose"
      Height          =   255
      Left            =   2640
      TabIndex        =   5
      Top             =   3480
      Width           =   2415
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "CheckBox mode, GunMetal Grey"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   3480
      Width           =   2415
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private TestArray() As String

Private Sub Check1_Click()
   Dim i As Long
   If Check1.Value = vbChecked Then
      For i = 1 To 4
         ucMorphListBox1(i).Enabled = True
      Next
   Else
      For i = 1 To 4
         ucMorphListBox1(i).Enabled = False
      Next
   End If
End Sub

Private Sub chkShowItemRect_Click()
   Dim i As Long
   For i = 1 To 4
      ucMorphListBox1(i).ShowSelectRect = (chkShowItemRect.Value = vbChecked)
   Next i
End Sub

Private Sub cmdAddItem_Click()
   Dim idx As Long, i As Long
   txtAddItem.Text = UCase(Trim(txtAddItem.Text))
   If txtAddItem.Text <> "" Then
      idx = Val(txtAddIndex.Text)
      For i = 1 To 4
         ucMorphListBox1(i).AddItem txtAddItem.Text, idx
         If optStats(i).Value = True Then
            DisplayPropertyStats i
         End If
      Next i
      txtAddItem.Text = ""
      txtAddIndex.Text = "-1"
   End If
End Sub

Private Sub cmdClear_Click()
   Dim i As Long
   For i = 1 To 4
      ucMorphListBox1(i).Clear
      If optStats(i).Value = True Then
         DisplayPropertyStats i
      End If
   Next i
End Sub

Private Sub cmdDisplayFrom_Click()
   Dim i As Long
   If Trim(txtDisplayFrom.Text) = "" Then
      Exit Sub
   End If
   For i = 1 To 4
      ucMorphListBox1(i).DisplayFrom Val(txtDisplayFrom.Text)
   Next i
End Sub

Private Sub cmdExit_Click()
   Unload Me
End Sub

Private Sub cmdFind_Click()

   If txtFind.Text <> "" Then
     If chkCase.Value = vbUnchecked Then
         lblFind.Caption = CStr(ucMorphListBox1(4).FindIndex(txtFind.Text))
      Else
         lblFind.Caption = CStr(ucMorphListBox1(4).FindIndex(txtFind.Text, True))
      End If
   End If

End Sub

Private Sub cmdLoad_Click()
   cmdClear_Click
   GenerateRandomList
End Sub

Private Sub Command1_Click()

End Sub

Private Sub Command2_Click()
   Dim i As Long
   If txtRemoveItem.Text <> "" And ucMorphListBox1(1).ListCount > 0 Then
      For i = 1 To 4
         ucMorphListBox1(i).RemoveItem Val(txtRemoveItem.Text)
         If optStats(i).Value = True Then
            DisplayPropertyStats CLng(i)
         End If
      Next i
   End If
End Sub

Private Sub Form_Load()

   Dim i As Long

   Randomize
   Form2.Show
   Form3.Show
   Form4.Show
   Form5.Show
   Form6.Show
   Form7.Show
   Form8.Show

End Sub

Private Sub GenerateRandomList()

   Dim Stopwatch As New CStopWatch
   Dim Max As Long, i As Long, j As Long, k As Long

   ReDim TestArray(1 To Val(txtAmount.Text))
   Max = Val(txtAmount.Text)

'  generate a list of 'max' random strings
   For i = 1 To Max
      k = Int(Rnd * 18) + 1
      s = ""
      For j = 1 To k '30
         letter = Int((90 - 65 + 1) * Rnd + 65)
         s = s & Chr(letter)
      Next j
      TestArray(i) = s
   Next i

   Stopwatch.Reset

   For i = 1 To 4

      ucMorphListBox1(i).RedrawFlag = False
      For j = 1 To Max
         ucMorphListBox1(i).AddItem TestArray(j)
      Next j
      ucMorphListBox1(i).RedrawFlag = True

      If optStats(i).Value = True Then
         DisplayPropertyStats i
      End If

   Next i

   lblUCLoadTime.Caption = "Avg. Time " & CStr(Max) & " Items: " & CStr((Stopwatch.Elapsed / 4) / 1000) & " secs"

End Sub

Private Sub optStats_Click(Index As Integer)
   DisplayPropertyStats CLng(Index)
End Sub

Private Sub DisplayPropertyStats(Index As Long)

   Dim lCount As Long, i As Long

   lblListCount.Caption = CStr(ucMorphListBox1(Index).ListCount)
   lblSelCount.Caption = CStr(ucMorphListBox1(Index).SelCount)
   lblListIndex.Caption = CStr(ucMorphListBox1(Index).ListIndex)
   lblTopIndex.Caption = CStr(ucMorphListBox1(Index).TopIndex) & ": " & _
                         ucMorphListBox1(Index).List(ucMorphListBox1(Index).TopIndex)

Exit Sub

'  for development verification purposes.
   lCount = 0
   For i = 0 To ucMorphListBox1(Index).ListCount - 1
      If ucMorphListBox1(Index).Selected(i) Then
         lCount = lCount + 1
      End If
   Next i
   lblCountOK.Caption = CStr(lCount)

End Sub

Private Sub ucMorphListBox1_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
   If optStats(Index).Value = True Then
      DisplayPropertyStats CLng(Index)
   End If
End Sub

Private Sub ucMorphListBox1_MouseEnter(Index As Integer)
'  small example of custom MouseEnter event usage.
   'If Index = 4 Then
   '   Set ucMorphListBox1(4).Picture = LoadPicture(App.Path & "\bluerivets.bmp")
   'End If
End Sub

Private Sub ucMorphListBox1_MouseLeave(Index As Integer)
'  small example of custom MouseLeave event usage.
   'If Index = 4 Then
   '   Set ucMorphListBox1(4).Picture = LoadPicture(App.Path & "\bluerivets2.bmp")
   'End If
End Sub

Private Sub ucMorphListBox1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'  an example of using the .MouseOverIndex method function call.
   Dim idx As Long
   idx = ucMorphListBox1(Index).MouseOverIndex(Y)
   If idx >= -1 Then
      lblMouseOverItem.Caption = ucMorphListBox1(Index).List(idx)
   End If
End Sub

Private Sub ucMorphListBox1_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   If optStats(Index).Value = True Then
      DisplayPropertyStats CLng(Index)
   End If
End Sub
