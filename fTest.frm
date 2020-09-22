VERSION 5.00
Begin VB.Form fTest 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Demo"
   ClientHeight    =   5250
   ClientLeft      =   45
   ClientTop       =   465
   ClientWidth     =   7095
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   350
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   473
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cbTrackbarPosition 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   345
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   3615
      Width           =   1815
   End
   Begin VB.CheckBox chkLocked 
      Caption         =   "Locked"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   360
      TabIndex        =   9
      Top             =   4425
      Width           =   1170
   End
   Begin VB.ComboBox cbBackStyle 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   360
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   2325
      Width           =   1815
   End
   Begin VB.ComboBox cbStyle 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   360
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   2970
      Width           =   1815
   End
   Begin VB.ListBox lstEvents 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3390
      Left            =   2880
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   1260
      Width           =   3855
   End
   Begin Test.ucComboTrackbar ucComboTrackbar1 
      Height          =   330
      Left            =   360
      TabIndex        =   1
      Top             =   1260
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   582
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      RangeMax        =   1
      RangeMin        =   -1
      RangePrecision  =   1
      Style           =   3
   End
   Begin VB.CheckBox chkEnabled 
      Caption         =   "Enabled"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   360
      TabIndex        =   8
      Top             =   4080
      Value           =   1  'Checked
      Width           =   1170
   End
   Begin VB.Label lblTrackbarPosition 
      Caption         =   "TrackbarPosition:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   360
      TabIndex        =   6
      Top             =   3390
      Width           =   1365
   End
   Begin VB.Label lblBackStyle 
      Caption         =   "BackStyle:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   375
      TabIndex        =   2
      Top             =   2100
      Width           =   1425
   End
   Begin VB.Label lblStyle 
      Caption         =   "Style:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   375
      TabIndex        =   4
      Top             =   2745
      Width           =   1365
   End
   Begin VB.Label lblSample 
      Caption         =   "Sample:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   375
      TabIndex        =   0
      Top             =   1005
      Width           =   1320
   End
End
Attribute VB_Name = "fTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_lEventCount As Long

Private Sub Form_Load()
    
    With cbBackStyle
        .AddItem "0 - [bsSolidColor]"
        .AddItem "1 - [bsImage]"
        .ListIndex = 0
    End With
    
    With cbStyle
        .AddItem "0 - [sClassic]"
        .AddItem "1 - [sFlat]"
        .AddItem "2 - [sFlatMono]"
        .AddItem "3 - [sThemed]"
        .ListIndex = 0
    End With
    
    With cbTrackbarPosition
        .AddItem "0 - [tpWide]"
        .AddItem "1 - [tpCentered]"
        .ListIndex = 0
    End With
    
    Set ucComboTrackbar1.BackImage = LoadResPicture(101, vbResBitmap)
    Let ucComboTrackbar1.Text = Format$(0, ucComboTrackbar1.PrecisionFormat)
End Sub

Private Sub Form_Paint()
    
    Me.Line (0, 0)-(Me.ScaleWidth, 50), vbWhite, BF
    
    Me.CurrentX = 10
    Me.CurrentY = 8
    Me.Font.Size = 12
    Me.Font.Bold = True
    Me.Print "ucComboTrackbar simple demo"
    
    Me.CurrentX = 10
    Me.CurrentY = 30
    Me.Font.Size = 9
    Me.Font.Bold = False
    Me.Print "Not all properties are shown"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call mCCEx.SafeEnd
End Sub

'//

Private Sub cbBackStyle_Click()
    ucComboTrackbar1.BackStyle = cbBackStyle.ListIndex
End Sub

Private Sub cbStyle_Click()
    ucComboTrackbar1.Style = cbStyle.ListIndex
End Sub

Private Sub cbTrackbarPosition_Click()
    ucComboTrackbar1.TrackbarPosition = cbTrackbarPosition.ListIndex
End Sub

Private Sub chkEnabled_Click()
    ucComboTrackbar1.Enabled = CBool(chkEnabled)
End Sub

Private Sub chkLocked_Click()
    ucComboTrackbar1.Locked = CBool(chkLocked)
End Sub

'//

Private Sub ucComboTrackbar1_Change()
    Call pvAddEventString("Change [Value = " & ucComboTrackbar1.Value & "]")
End Sub

Private Sub ucComboTrackbar1_Scroll()
    Call pvAddEventString("Scroll")
End Sub

Private Sub ucComboTrackbar1_GotFocus()
    Call pvAddEventString("GotFocus")
End Sub

Private Sub ucComboTrackbar1_LostFocus()
    Call pvAddEventString("LostFocus")
End Sub

'Private Sub ucComboTrackbar1_KeyDown(KeyCode As Integer, Shift As Integer)
'    Call pvAddEventString("KeyDown (KeyCode: " & KeyCode & " Shift: " & Shift & ")")
'End Sub
'
'Private Sub ucComboTrackbar1_KeyPress(KeyAscii As Integer)
'    Call pvAddEventString("KeyPress (KeyAscii: " & KeyAscii & ")")
'End Sub
'
'Private Sub ucComboTrackbar1_KeyUp(KeyCode As Integer, Shift As Integer)
'    Call pvAddEventString("KeyUp (KeyCode: " & KeyCode & " Shift: " & Shift & ")")
'End Sub

Private Sub ucComboTrackbar1_ThemeChanged()
    Call pvAddEventString("ThemeChanged")
End Sub

Private Sub ucComboTrackbar1_TrackbarShow()
    Call pvAddEventString("TrackbarShow")
End Sub

Private Sub ucComboTrackbar1_TrackbarHide()
    Call pvAddEventString("TrackbarHide")
End Sub

Private Sub pvAddEventString(ByVal sString As String)
    
    With lstEvents
        m_lEventCount = m_lEventCount + 1
        
        If (.ListCount = 16) Then
            Call .RemoveItem(0)
        End If
        Call .AddItem(Format$(m_lEventCount, "00000 ") & sString)
        .ListIndex = .ListCount - 1
    End With
End Sub
