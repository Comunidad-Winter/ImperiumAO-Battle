VERSION 5.00
Begin VB.Form frmMacro 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Asignar Acci�n"
   ClientHeight    =   4110
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   3000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4110
   ScaleWidth      =   3000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   3600
      Width           =   1335
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   1560
      TabIndex        =   6
      Top             =   3600
      Width           =   1335
   End
   Begin VB.TextBox txtComando 
      Height          =   285
      Left            =   480
      TabIndex        =   5
      Top             =   2250
      Width           =   2415
   End
   Begin VB.OptionButton optEquipar 
      Caption         =   "Equipar item elegido"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   3180
      Width           =   2775
   End
   Begin VB.OptionButton optUsar 
      Caption         =   "Usar item elegido"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   2880
      Width           =   2775
   End
   Begin VB.OptionButton optHechi 
      Caption         =   "Usar hechizo elegido"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   2580
      Width           =   2775
   End
   Begin VB.OptionButton optComando 
      Caption         =   "Comando"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   1920
      Width           =   2775
   End
   Begin VB.Line Line2 
      X1              =   150
      X2              =   2850
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Line Line1 
      X1              =   150
      X2              =   2850
      Y1              =   1830
      Y2              =   1830
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmMacro.frx":0000
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1545
      Left            =   210
      TabIndex        =   9
      Top             =   390
      Width           =   2625
   End
   Begin VB.Label numF 
      Height          =   135
      Left            =   1830
      TabIndex        =   8
      Top             =   4860
      Width           =   855
   End
   Begin VB.Label lblF 
      Alignment       =   2  'Center
      Caption         =   "Macro F1"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   -30
      Width           =   3015
   End
End
Attribute VB_Name = "frmMacro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
txtComando.Text = ""
Me.Hide
End Sub

Private Sub cmdOk_Click()

If optComando.value = True And txtComando.Text = "" Then
    MsgBox "Ingrese un comando"
    Exit Sub
End If

Macros(Val(numF - 1)).comando = ""
Macros(Val(numF - 1)).Equipar = 0
Macros(Val(numF - 1)).Hechizo = 0
Macros(Val(numF - 1)).Usar = 0

If optComando.value Then
    Macros(Val(numF - 1)).comando = Trim(txtComando.Text)
ElseIf optHechi.value Then
    Macros(Val(numF - 1)).Hechizo = frmMain.hlst.listIndex + 1
ElseIf optUsar.value Then
    Macros(Val(numF - 1)).Usar = Inventario.SelectedItem
ElseIf optEquipar.value Then
    Macros(Val(numF - 1)).Equipar = Inventario.SelectedItem
Else
    MsgBox "Seleccione alguna opci�n"
    Exit Sub
End If

Me.Hide
GuardarMacros
CargarMacros (False)
End Sub

