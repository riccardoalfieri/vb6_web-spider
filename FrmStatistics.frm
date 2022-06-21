VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmStatistics 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "WebSpider :: Statistics"
   ClientHeight    =   9180
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11745
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9180
   ScaleWidth      =   11745
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox GFXFrameX3 
      Height          =   855
      Left            =   360
      ScaleHeight     =   795
      ScaleWidth      =   5235
      TabIndex        =   17
      Top             =   7680
      Width           =   5295
   End
   Begin VB.PictureBox GFXFrameX1 
      Height          =   1455
      Left            =   6360
      ScaleHeight     =   1395
      ScaleWidth      =   4635
      TabIndex        =   15
      Top             =   7680
      Width           =   4695
      Begin VB.PictureBox GFXFrameX2 
         Height          =   615
         Left            =   240
         ScaleHeight     =   555
         ScaleWidth      =   2595
         TabIndex        =   16
         Top             =   600
         Width           =   2655
      End
   End
   Begin VB.PictureBox Picture2 
      Height          =   4455
      Left            =   6720
      Picture         =   "FrmStatistics.frx":0000
      ScaleHeight     =   4395
      ScaleWidth      =   4755
      TabIndex        =   14
      Top             =   0
      Width           =   4815
   End
   Begin VB.CommandButton CmdToggle 
      Caption         =   "Début Spider"
      Height          =   375
      Left            =   5280
      TabIndex        =   13
      Top             =   840
      Width           =   1215
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   2520
      TabIndex        =   12
      Top             =   1920
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin MSComctlLib.ListView LstURL 
      Height          =   4695
      Left            =   0
      TabIndex        =   0
      Top             =   2640
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   8281
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "URL"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Scanned"
         Object.Width           =   1588
      EndProperty
   End
   Begin MSComctlLib.ListView LstEmail 
      Height          =   2415
      Left            =   6600
      TabIndex        =   1
      Top             =   4920
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   4260
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Name"
         Object.Width           =   5468
      EndProperty
   End
   Begin VB.Label Label9 
      Caption         =   "Page en cours:"
      Height          =   255
      Left            =   2400
      TabIndex        =   11
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Label lblPage 
      Height          =   255
      Left            =   3840
      TabIndex        =   10
      Top             =   1680
      Width           =   2655
   End
   Begin VB.Label Label8 
      Caption         =   "E-Mail trouvées"
      Height          =   255
      Left            =   2400
      TabIndex        =   9
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Label lblEmail 
      Caption         =   "0"
      Height          =   255
      Left            =   3840
      TabIndex        =   8
      Top             =   1320
      Width           =   2655
   End
   Begin VB.Label Label7 
      Caption         =   "URL trouvées:"
      Height          =   255
      Left            =   2400
      TabIndex        =   7
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label lblUrl 
      Caption         =   "0"
      Height          =   255
      Left            =   3840
      TabIndex        =   6
      Top             =   1080
      Width           =   2655
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "Pages numérisées:"
      Height          =   255
      Left            =   2400
      TabIndex        =   5
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label lblScanned 
      Caption         =   "0"
      Height          =   255
      Left            =   3840
      TabIndex        =   4
      Top             =   840
      Width           =   2655
   End
   Begin VB.Line Line4 
      BorderWidth     =   2
      X1              =   2520
      X2              =   6600
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00808080&
      X1              =   2535
      X2              =   6615
      Y1              =   2295
      Y2              =   2295
   End
   Begin VB.Label Label4 
      Caption         =   "Adresses E-Mail"
      Height          =   255
      Left            =   6600
      TabIndex        =   3
      Top             =   4680
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "URL's:"
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      X1              =   2535
      X2              =   6615
      Y1              =   735
      Y2              =   735
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   2520
      X2              =   6600
      Y1              =   720
      Y2              =   720
   End
End
Attribute VB_Name = "FrmStatistics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdToggle_Click()
    LogURL FrmMain.TxtUrl
    Toggle
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Cancel = 1
    FrmStatistics.Visible = False
    FrmMain.Visible = True
  
End Sub
