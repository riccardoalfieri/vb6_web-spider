VERSION 5.00
Begin VB.Form MessagePopUp 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1800
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   4395
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1800
   ScaleWidth      =   4395
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox GFXFrameX3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   960
      ScaleHeight     =   315
      ScaleWidth      =   3195
      TabIndex        =   8
      Top             =   1200
      Width           =   3255
   End
   Begin VB.PictureBox GFXFrameX2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   800
      ScaleHeight     =   285
      ScaleWidth      =   3315
      TabIndex        =   7
      Top             =   700
      Width           =   3375
   End
   Begin VB.PictureBox GFXFrameX1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   960
      ScaleHeight     =   315
      ScaleWidth      =   3195
      TabIndex        =   6
      Top             =   360
      Width           =   3255
   End
   Begin VB.Timer PopDelay 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   840
      Top             =   1320
   End
   Begin VB.OptionButton PopDown 
      Caption         =   "Option1"
      Height          =   195
      Left            =   240
      TabIndex        =   2
      Top             =   1560
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.OptionButton PopUp 
      Caption         =   "Option1"
      Height          =   195
      Left            =   0
      TabIndex        =   1
      Top             =   1560
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Timer MessageUpDown 
      Interval        =   1
      Left            =   480
      Top             =   1320
   End
   Begin VB.TextBox txtTopPosition 
      Height          =   285
      Left            =   0
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   1080
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lDescription 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   5
      Top             =   1200
      Width           =   4095
      WordWrap        =   -1  'True
   End
   Begin VB.Label lTitle 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Title"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   840
      TabIndex        =   4
      Top             =   360
      Width           =   3255
   End
   Begin VB.Image Image3 
      Height          =   480
      Left            =   120
      Picture         =   "MessagePopUp.frx":0000
      Top             =   360
      Width           =   480
   End
   Begin VB.Label lClose 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   3600
      TabIndex        =   3
      Top             =   30
      Width           =   615
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FF8080&
      BackStyle       =   1  'Opaque
      Height          =   255
      Left            =   0
      Top             =   0
      Width           =   4455
   End
End
Attribute VB_Name = "MessagePopUp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'MESSAGE ON TOP
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, _
    ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, _
    ByVal cy As Long, ByVal wFlags As Long) As Long

'GET SCREEN RIGHT CORNER
Private Type RECT
    Left As Long
    top As Long
    Right As Long
    Bottom As Long
End Type

Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, _
    ByVal uParam As Long, ByRef lpvParam As RECT, ByVal fuWinIni As Long) As Long

Private Const SPI_GETWORKAREA = 48

'FORM HEIGHT CALCULATION
Dim popuptop As Long
Dim popupdown As Long

'MESSAGE ON TOP
Public Sub MessageOnTop(hWindow As Long, bTopMost As Boolean)
    
Const SWP_NOSIZE = &H1
Const SWP_NOMOVE = &H2
Const SWP_NOACTIVATE = &H10
Const SWP_SHOWWINDOW = &H40
Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2
    
Dim wFlags
Dim placement
    
wFlags = SWP_NOMOVE Or SWP_NOSIZE Or SWP_SHOWWINDOW Or SWP_NOACTIVATE
    
placement = HWND_TOPMOST
    
SetWindowPos hWindow, placement, 0, 0, 0, 0, wFlags

End Sub

'POSITION THE FORM IN THE RIGHT CORNER OF SCREEN
Private Sub PlaceMessageInLowerRight(ByVal frm As Form, ByVal right_margin As Single, ByVal bottom_margin As Single)

Dim wa_info As RECT

'If MainForm.Taskbar.Value = True Then

    If SystemParametersInfo(SPI_GETWORKAREA, 0, wa_info, 0) <> 0 Then

        'GOT POSITION, PLACE THE FORM NOW
        frm.Left = ScaleX(wa_info.Right, vbPixels, vbTwips) - Width - right_margin
        frm.top = ScaleY(wa_info.Bottom, vbPixels, vbTwips) - Height - bottom_margin
    
        txtTopPosition.Text = frm.top 'THIS IS THE MESSAGE'S LEFT (BOTTOM) TOP VALUE WHERE THE FORM SHOULD STOP
        'LAST DIGIT OF THIS VALUE SHOULD BE ZERO TO PREVENT LENGHT ERROR
        txtTopPosition.Text = Replace(txtTopPosition, Right(txtTopPosition, 1), 0)
        
    End If

'End If


End Sub

Private Sub Command1_Click()

FrmInternetBrowser.WindowState = vbMaximized
End Sub

Private Sub Form_Load()
'GFXFrameX2.Caption = "Voulez un design personnalisé? "
'GFXFrameX3.Caption = "CLIQUEZ ICI "


Call MessageOnTop(Me.hwnd, True) 'MESSAGE ON TOP

PlaceMessageInLowerRight Me, 0, 0 'MESSAGE PLACEMENT

PopUp.Value = True 'MESSAGE POPUP

Me.top = Screen.Height 'PLACE MESSAGE IN THE BOTTOM OF THE SCREEN TO CALL UP

End Sub

Private Sub GFXFrameX3_GotFocus()
Call Shell("explorer http://axasoft.altervista.org/joomla25/fr/contactez-moi", vbMaximizedFocus)
End Sub

Private Sub lClose_Click()

PopDelay.Enabled = False  'DONT ALLOW TO PROCEED TO DELAY TIME
MessageUpDown.Enabled = True   'KEEP MESSAGE MOVING
PopDown.Value = True    'ALLOW MESSAGE DOWN
PopUp.Value = False     'SKIP MESSAGE UP

End Sub

Private Sub lClose_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

lClose.ForeColor = vbWhite

End Sub

Private Sub lClose_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

lClose.ForeColor = vbBlack

End Sub





Private Sub MessageUpDown_Timer()
    
'If PopUp.Value = True Then  'MESSAGE POPUP
    
    popuptop = Me.top - 10 'GRADUALLY START THE MESSAGE UP (WARNING: CHANGING THIS VALUE MAY NOT WORK)
    Me.top = popuptop
    
    'Debug.Print Me.Top
                                
        If Me.top = txtTopPosition.Text Then    'THIS IS THE VALUE OF MESSAGE TOP WHEN IT FULLY APPEARES
            
            PopUp.Value = False     'STOP MESSAGE UP
            MessageUpDown.Enabled = False  'DISABLE MESSAGE POPUP
            PopDelay.Enabled = True   'START DELAY TIME
                
        End If
    
'End If
    

End Sub

Private Sub PopDelay_Timer()

PopDelay.Interval = PopDelay.Interval + 1000    'DELAY COUNTER

If PopDelay.Interval = 4000 Then  'THIS IS DELAY TIME UNTIL WHEN THE MESSAGE HAS TO BE ON THE SCREEN.

    PopDown.Value = True    'ENABLE POPDOWN
    MessageUpDown.Enabled = True   'ENABLE FORM MOVEMENT
    PopDelay.Enabled = False  'DISABLE DELAY COUNTER

End If

End Sub
