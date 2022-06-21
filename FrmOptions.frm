VERSION 5.00
Begin VB.Form FrmOptions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "WebSpider :: Options"
   ClientHeight    =   4920
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4920
   ScaleWidth      =   11880
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "URL's trouvés"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8880
      TabIndex        =   9
      Top             =   4080
      Width           =   2415
   End
   Begin VB.CommandButton Command2 
      Caption         =   "e-mails trouvés"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6960
      TabIndex        =   8
      Top             =   4080
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Pages numérisées"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4680
      TabIndex        =   7
      Top             =   4080
      Width           =   2295
   End
   Begin VB.TextBox txtSave 
      Height          =   285
      Left            =   6000
      TabIndex        =   6
      Text            =   "C:\Spider.[name].log"
      Top             =   2280
      Width           =   3255
   End
   Begin VB.CheckBox ChkSave 
      Caption         =   "Enregistrer trouver à:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5760
      TabIndex        =   5
      Top             =   1920
      Value           =   1  'Checked
      Width           =   3375
   End
   Begin VB.OptionButton OptNothing 
      Caption         =   "Ne rien faire"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5760
      TabIndex        =   4
      Top             =   1560
      Width           =   3975
   End
   Begin VB.TextBox txtSearch 
      Height          =   285
      Left            =   6960
      TabIndex        =   3
      Top             =   1150
      Width           =   3975
   End
   Begin VB.OptionButton OptFind 
      Caption         =   "Trouver:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5760
      TabIndex        =   2
      Top             =   1200
      Width           =   3975
   End
   Begin VB.OptionButton OptEmail 
      Caption         =   "Trouver des adresses e-mail"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5760
      TabIndex        =   1
      Top             =   840
      Value           =   -1  'True
      Width           =   3855
   End
   Begin VB.PictureBox Picture1 
      Height          =   4455
      Left            =   120
      Picture         =   "FrmOptions.frx":0000
      ScaleHeight     =   4395
      ScaleWidth      =   4515
      TabIndex        =   0
      Top             =   360
      Width           =   4575
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      X1              =   6255
      X2              =   10335
      Y1              =   735
      Y2              =   735
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   6240
      X2              =   10320
      Y1              =   720
      Y2              =   720
   End
End
Attribute VB_Name = "FrmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Shell "c:\windows\system32\notepad.exe spider.search.log", vbNormalFocus

End Sub

Private Sub Command2_Click()
Shell "c:\windows\system32\notepad.exe spider.email.log", vbNormalFocus

End Sub

Private Sub Command3_Click()
Shell "c:\windows\system32\notepad.exe spider.url.log", vbNormalFocus

End Sub

Private Sub Form_Load()
    txtSave = App.Path & "\Spider.[name].log"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Cancel = 1
    FrmOptions.Visible = False
    FrmMain.Visible = True
End Sub

Private Sub OptEmail_Click()
    FrmStatistics.LstEmail.Visible = True
    FrmStatistics.Label4.Visible = True
    FrmStatistics.Label4.Caption = "E-mail Adresse:"
    FrmStatistics.LstEmail.ColumnHeaders(1).Text = "Adresse"
    FrmStatistics.LstEmail.ListItems.Clear
    FrmStatistics.Label8.Caption = "Email Trouvé: "
End Sub

Private Sub OptFind_Click()
    FrmStatistics.LstEmail.Visible = True
    FrmStatistics.Label4.Visible = True
    FrmStatistics.Label4.Caption = "Résultat de la recherche:"
    FrmStatistics.LstEmail.ColumnHeaders(1).Text = "URL"
    FrmStatistics.LstEmail.ListItems.Clear
    FrmStatistics.Label8.Caption = "Trouvé: "
End Sub

Private Sub OptNothing_Click()
    FrmStatistics.LstEmail.Visible = False
    FrmStatistics.Label4.Visible = False
End Sub
