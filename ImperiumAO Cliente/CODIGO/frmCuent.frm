VERSION 5.00
Begin VB.Form frmCuent 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11970
   BeginProperty Font 
      Name            =   "Georgia"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MouseIcon       =   "frmCuent.frx":0000
   MousePointer    =   99  'Custom
   Picture         =   "frmCuent.frx":0CCA
   ScaleHeight     =   9000
   ScaleWidth      =   11970
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox PJ 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Index           =   9
      Left            =   8610
      MouseIcon       =   "frmCuent.frx":6B201
      MousePointer    =   99  'Custom
      Picture         =   "frmCuent.frx":6BECB
      ScaleHeight     =   1215
      ScaleWidth      =   735
      TabIndex        =   35
      Top             =   5550
      Width           =   735
   End
   Begin VB.PictureBox PJ 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Index           =   8
      Left            =   8580
      MouseIcon       =   "frmCuent.frx":6C166
      MousePointer    =   99  'Custom
      Picture         =   "frmCuent.frx":6CE30
      ScaleHeight     =   1215
      ScaleWidth      =   735
      TabIndex        =   34
      Top             =   3720
      Width           =   735
   End
   Begin VB.PictureBox PJ 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Index           =   5
      Left            =   4080
      MouseIcon       =   "frmCuent.frx":6D0CB
      MousePointer    =   99  'Custom
      Picture         =   "frmCuent.frx":6DD95
      ScaleHeight     =   1215
      ScaleWidth      =   735
      TabIndex        =   32
      Top             =   5520
      Width           =   735
   End
   Begin VB.PictureBox PJ 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Index           =   6
      Left            =   5550
      MouseIcon       =   "frmCuent.frx":6E030
      MousePointer    =   99  'Custom
      Picture         =   "frmCuent.frx":6ECFA
      ScaleHeight     =   1215
      ScaleWidth      =   735
      TabIndex        =   31
      Top             =   5520
      Width           =   735
   End
   Begin VB.PictureBox PJ 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Index           =   7
      Left            =   7110
      MouseIcon       =   "frmCuent.frx":6EF95
      MousePointer    =   99  'Custom
      Picture         =   "frmCuent.frx":6FC5F
      ScaleHeight     =   1215
      ScaleWidth      =   735
      TabIndex        =   30
      Top             =   5520
      Width           =   735
   End
   Begin VB.PictureBox PJ 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Index           =   4
      Left            =   2640
      MouseIcon       =   "frmCuent.frx":6FEFA
      MousePointer    =   99  'Custom
      Picture         =   "frmCuent.frx":70BC4
      ScaleHeight     =   1215
      ScaleWidth      =   735
      TabIndex        =   4
      Top             =   5520
      Width           =   735
   End
   Begin VB.PictureBox PJ 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Index           =   3
      Left            =   7110
      MouseIcon       =   "frmCuent.frx":70E5F
      MousePointer    =   99  'Custom
      Picture         =   "frmCuent.frx":71B29
      ScaleHeight     =   1215
      ScaleWidth      =   735
      TabIndex        =   3
      Top             =   3720
      Width           =   735
   End
   Begin VB.PictureBox PJ 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Index           =   2
      Left            =   5610
      MouseIcon       =   "frmCuent.frx":71DC4
      MousePointer    =   99  'Custom
      Picture         =   "frmCuent.frx":72A8E
      ScaleHeight     =   1215
      ScaleWidth      =   735
      TabIndex        =   2
      Top             =   3720
      Width           =   735
   End
   Begin VB.PictureBox PJ 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Index           =   1
      Left            =   4110
      MouseIcon       =   "frmCuent.frx":72D29
      MousePointer    =   99  'Custom
      Picture         =   "frmCuent.frx":739F3
      ScaleHeight     =   1215
      ScaleWidth      =   735
      TabIndex        =   1
      Top             =   3750
      Width           =   735
   End
   Begin VB.PictureBox PJ 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillColor       =   &H0000FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Index           =   0
      Left            =   2640
      MouseIcon       =   "frmCuent.frx":73C8E
      MousePointer    =   99  'Custom
      Picture         =   "frmCuent.frx":74958
      ScaleHeight     =   1215
      ScaleWidth      =   735
      TabIndex        =   0
      Top             =   3720
      Width           =   735
   End
   Begin VB.Image Image6 
      Height          =   285
      Left            =   11340
      Top             =   60
      Width           =   285
   End
   Begin VB.Image Image5 
      Height          =   315
      Left            =   11670
      Top             =   60
      Width           =   285
   End
   Begin VB.Image Image4 
      Height          =   555
      Left            =   8010
      Top             =   2100
      Width           =   1725
   End
   Begin VB.Image Image3 
      Height          =   585
      Left            =   4170
      Top             =   7590
      Width           =   1725
   End
   Begin VB.Image Image2 
      Height          =   615
      Left            =   2250
      Top             =   7560
      Width           =   1755
   End
   Begin VB.Image Image1 
      Height          =   645
      Left            =   8010
      Top             =   7560
      Width           =   1785
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   9
      Left            =   8010
      TabIndex        =   41
      Top             =   6870
      Width           =   1935
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   9
      Left            =   8250
      TabIndex        =   40
      Top             =   6720
      Width           =   1455
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   8
      Left            =   8010
      TabIndex        =   39
      Top             =   5040
      Width           =   1935
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   8
      Left            =   8250
      TabIndex        =   38
      Top             =   4890
      Width           =   1455
   End
   Begin VB.Label nombre 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Nada"
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   9
      Left            =   8100
      TabIndex        =   37
      Top             =   5310
      Width           =   1815
   End
   Begin VB.Label nombre 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Nada"
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   8
      Left            =   8100
      TabIndex        =   36
      Top             =   3510
      Width           =   1815
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre de la Cuenta:"
      ForeColor       =   &H80000007&
      Height          =   255
      Left            =   720
      TabIndex        =   33
      Top             =   480
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   255
      Index           =   7
      Left            =   6510
      TabIndex        =   29
      Top             =   6870
      Width           =   1935
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   255
      Index           =   6
      Left            =   4950
      TabIndex        =   28
      Top             =   6870
      Width           =   1935
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   255
      Index           =   5
      Left            =   3480
      TabIndex        =   27
      Top             =   6870
      Width           =   1935
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   7
      Left            =   6750
      TabIndex        =   26
      Top             =   6690
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   6
      Left            =   5190
      TabIndex        =   25
      Top             =   6690
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   5
      Left            =   3720
      TabIndex        =   24
      Top             =   6690
      Width           =   1455
   End
   Begin VB.Label nombre 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Nada"
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   7
      Left            =   6570
      TabIndex        =   23
      Top             =   5310
      Width           =   1815
   End
   Begin VB.Label nombre 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Nada"
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   6
      Left            =   5010
      TabIndex        =   22
      Top             =   5310
      Width           =   1815
   End
   Begin VB.Label nombre 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Nada"
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   5
      Left            =   3570
      TabIndex        =   21
      Top             =   5310
      Width           =   1815
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "PJClick"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2340
      TabIndex        =   20
      Top             =   2490
      Width           =   2895
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   4
      Left            =   2250
      TabIndex        =   19
      Top             =   6720
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   6690
      TabIndex        =   18
      Top             =   4890
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   5190
      TabIndex        =   17
      Top             =   4920
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   3690
      TabIndex        =   16
      Top             =   4890
      Width           =   1455
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   4
      Left            =   2010
      TabIndex        =   15
      Top             =   6870
      Width           =   1935
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   255
      Index           =   3
      Left            =   6450
      TabIndex        =   14
      Top             =   5040
      Width           =   1935
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   255
      Index           =   2
      Left            =   4950
      TabIndex        =   13
      Top             =   5070
      Width           =   1935
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   255
      Index           =   1
      Left            =   3450
      TabIndex        =   12
      Top             =   5070
      Width           =   1935
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   2010
      TabIndex        =   11
      Top             =   5070
      Width           =   1935
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   2250
      TabIndex        =   10
      Top             =   4920
      Width           =   1455
   End
   Begin VB.Label nombre 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Nada"
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   4
      Left            =   2100
      TabIndex        =   9
      Top             =   5310
      Width           =   1815
   End
   Begin VB.Label nombre 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Nada"
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   3
      Left            =   6570
      TabIndex        =   8
      Top             =   3510
      Width           =   1815
   End
   Begin VB.Label nombre 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Nada"
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   2
      Left            =   5100
      TabIndex        =   7
      Top             =   3510
      Width           =   1815
   End
   Begin VB.Label nombre 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Nada"
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   1
      Left            =   3600
      TabIndex        =   6
      Top             =   3510
      Width           =   1815
   End
   Begin VB.Label nombre 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Nada"
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   0
      Left            =   2100
      TabIndex        =   5
      Top             =   3480
      Width           =   1815
   End
End
Attribute VB_Name = "frmCuent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private Sub Command1_Click()
If PJClickeado = "Nada" Then
MsgBox "Seleccione un pj"
End If
Call Audio.PlayWave(SND_CLICK)
UserName = PJClickeado
SendData ("OOLOGI" & UserName)
Unload Me
End Sub

Private Sub Command4_Click()
frmBorrar.Show , frmCuent
End Sub

Private Sub Form_load()
Dim i As Integer
Label3.Caption = UserName
End Sub
Private Sub Command2_Click()
frmMain.Socket1.Disconnect
Unload Me
frmConnect.Show
End Sub

Private Sub Command3_Click()
Call Audio.PlayWave(SND_CLICK)

If Nombre(9).Caption <> "Nada" Then
    MsgBox "Tu cuenta ha llegado al máximo de personajes."
    Exit Sub
End If

    EstadoLogin = Dados
    frmCrearPersonaje.Show vbModal
    Me.MousePointer = 11
    
End Sub

Private Sub Image1_Click()
If PJClickeado = "Nada" Then
MsgBox "Seleccione un pj"
End If
Call Audio.PlayWave(SND_CLICK)
UserName = PJClickeado
SendData ("OOLOGI" & UserName)
Unload Me
End Sub

Private Sub Image2_Click()
Call Audio.PlayWave(SND_CLICK)

If Nombre(9).Caption <> "Nada" Then
    MsgBox "Tu cuenta ha llegado al máximo de personajes."
    Exit Sub
End If

    EstadoLogin = Dados
    frmCrearPersonaje.Show vbModal
    Me.MousePointer = 11
    
End Sub

Private Sub Image3_Click()
frmBorrar.Show , frmCuent
End Sub

Private Sub Image4_Click()
frmMain.Socket1.Disconnect
Unload Me
frmConnect.Show
End Sub

Private Sub Image5_Click()
End
End Sub

Private Sub Image6_Click()
Me.WindowState = vbMinimized
End Sub

Private Sub nombre_dblClick(index As Integer)
If PJClickeado = "Nada" Then Exit Sub
Call Audio.PlayWave(SND_CLICK)
UserName = PJClickeado
SendData ("OOLOGI" & UserName)
Unload Me
End Sub
Private Sub nombre_Click(index As Integer)
PJClickeado = frmCuent.Nombre(index).Caption
End Sub
Private Sub PJ_Click(index As Integer)
PJClickeado = frmCuent.Nombre(index).Caption
End Sub

Private Sub PJ_dblClick(index As Integer)
If PJClickeado = "Nada" Then Exit Sub
Call Audio.PlayWave(SND_CLICK)
UserName = PJClickeado
SendData ("OOLOGI" & UserName)
Unload Me
End Sub


