VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form FrmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "WebSpider"
   ClientHeight    =   8325
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14790
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8325
   ScaleWidth      =   14790
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "Arrêter"
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
      Left            =   11400
      TabIndex        =   10
      Top             =   6240
      Width           =   1335
   End
   Begin VB.PictureBox Picture2 
      Height          =   4455
      Left            =   240
      Picture         =   "FrmMain.frx":0000
      ScaleHeight     =   4395
      ScaleWidth      =   4755
      TabIndex        =   9
      Top             =   840
      Width           =   4815
   End
   Begin VB.CommandButton CmdOptions 
      Caption         =   "Options"
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
      Left            =   10200
      TabIndex        =   8
      Top             =   6240
      Width           =   1215
   End
   Begin VB.CommandButton CmdView 
      Caption         =   "Afficher les statistiques"
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
      Left            =   7680
      TabIndex        =   7
      Top             =   6240
      Width           =   2535
   End
   Begin VB.CommandButton CmdToggle 
      Caption         =   "Début Spider"
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
      Left            =   6000
      TabIndex        =   6
      Top             =   6240
      Width           =   1695
   End
   Begin VB.TextBox txtMemory 
      Height          =   285
      Left            =   8520
      TabIndex        =   3
      Text            =   "-1"
      Top             =   5400
      Width           =   3735
   End
   Begin VB.TextBox TxtUrl 
      Height          =   285
      Left            =   7320
      TabIndex        =   1
      Text            =   "http://www.axasoft.altervista.org/go.html"
      Top             =   5040
      Width           =   4935
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   6000
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.PictureBox GFXFrameX1 
      CausesValidation=   0   'False
      Height          =   4095
      Left            =   6720
      ScaleHeight     =   4035
      ScaleWidth      =   5955
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   240
      Width           =   6015
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
         Height          =   615
         Index           =   0
         Left            =   240
         ScaleHeight     =   555
         ScaleWidth      =   4995
         TabIndex        =   12
         Top             =   480
         Width           =   5055
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
         Height          =   615
         Index           =   1
         Left            =   240
         ScaleHeight     =   555
         ScaleWidth      =   4995
         TabIndex        =   13
         Top             =   1320
         Width           =   5055
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
         Height          =   615
         Index           =   2
         Left            =   240
         ScaleHeight     =   555
         ScaleWidth      =   4995
         TabIndex        =   14
         Top             =   2160
         Width           =   5055
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
         Height          =   615
         Index           =   3
         Left            =   240
         ScaleHeight     =   555
         ScaleWidth      =   4995
         TabIndex        =   15
         Top             =   3000
         Width           =   5055
      End
   End
   Begin VB.Label lblScanned 
      Caption         =   "0"
      Height          =   255
      Left            =   7920
      TabIndex        =   5
      Top             =   5880
      Width           =   2655
   End
   Begin VB.Label Label5 
      Caption         =   "Pages numérisées:"
      Height          =   255
      Left            =   6000
      TabIndex        =   4
      Top             =   5880
      Width           =   1815
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00808080&
      X1              =   6495
      X2              =   10575
      Y1              =   5775
      Y2              =   5775
   End
   Begin VB.Line Line3 
      BorderWidth     =   2
      X1              =   6000
      X2              =   12120
      Y1              =   5760
      Y2              =   5760
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "URL De tenir en mémoire:"
      Height          =   255
      Left            =   6000
      TabIndex        =   2
      Top             =   5430
      Width           =   2415
   End
   Begin VB.Label Label3 
      Caption         =   "URL de base"
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
      Left            =   6000
      TabIndex        =   0
      Top             =   5055
      Width           =   1455
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   6120
      X2              =   12120
      Y1              =   4920
      Y2              =   4920
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      X1              =   6495
      X2              =   10575
      Y1              =   4935
      Y2              =   4935
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Private Sub CmdOptions_Click()
    FrmOptions.Show
  '  Me.Visible = False
End Sub

Private Sub CmdToggle_Click()
Dim updatestring As String
On Error Resume Next



  
  



    Toggle
End Sub

Private Sub CmdView_Click()
    FrmStatistics.Show
    Me.Visible = False
End Sub


Private Sub Command3_Click()
End
End Sub

Private Sub Form_Load()
' apre i tre files dati per l'append

        Open Replace(FrmOptions.txtSave, "[name]", "search") For Append As #1
        Open Replace(FrmOptions.txtSave, "[name]", "email") For Append As #2
        Open Replace(FrmOptions.txtSave, "[name]", "url") For Append As #3
    
    
    
    FrmStatistics.Show
    FrmStatistics.Visible = False
    FrmOptions.Show
    FrmOptions.Visible = False
    
    SpiderOnline = False
   ' GFXFrameX1.Caption = "LA TOILE DE L'ARAIGNEE"
   ' GFXFrameX2(0).Caption = "Trouver singles sur facebook"
    '  GFXFrameX2(1).Caption = "Recherche de clients et fournisseurs en ligne"
 'GFXFrameX2(2).Caption = "Crée le trafic à votre site Web"
'GFXFrameX2(3).Caption = "Recherche et envoi d'e-mails"

End Sub

Private Sub Form_Unload(Cancel As Integer)
    If MsgBox("Êtes-vous sûr de vouloir quitter", vbYesNo) = vbYes Then
        Close #1
        Close #2
        Close #3
        End
    Else
        Cancel = 1
    End If
End Sub

         Public Sub JamieHTMLParser1_HTMLText(text As String, property As String, propertyValue As String)
  If Len(text) > 4 Then
  LogEmail text
  End If
  
     End Sub
              
          Private Sub JamieHTMLParser1_HTMLTagClose(Tag As String, text As String, property As String, propertyValue As String)
          End Sub
          Private Sub JamieHTMLParser1_HTMLTagBegin(Tag As String, text As String)

          End Sub
Private Sub JamieHTMLParser1_HTMLProperty(property As String, propertyValue As String, text As String, exPropertyValue As String)
End Sub


