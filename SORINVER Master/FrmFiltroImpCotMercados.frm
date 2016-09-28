VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmFiltroImpCotMercados 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Importación cotizaciones Mercados"
   ClientHeight    =   1560
   ClientLeft      =   4155
   ClientTop       =   3870
   ClientWidth     =   5910
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Trebuchet MS"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmFiltroImpCotMercados.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1560
   ScaleWidth      =   5910
   Begin VB.CommandButton BCancelar 
      Caption         =   "&Cancelar"
      Height          =   615
      Left            =   4320
      TabIndex        =   3
      Top             =   840
      Width           =   1455
   End
   Begin VB.CommandButton BAceptar 
      Caption         =   "&Aceptar"
      Height          =   615
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   1455
   End
   Begin MSComCtl2.DTPicker DTFechaDesde 
      Height          =   375
      Left            =   1560
      TabIndex        =   0
      Top             =   240
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CalendarBackColor=   -2147483639
      CalendarForeColor=   -2147483625
      CalendarTitleBackColor=   8388608
      CalendarTitleForeColor=   -2147483639
      CalendarTrailingForeColor=   -2147483626
      Format          =   50003969
      CurrentDate     =   36892
   End
   Begin MSComCtl2.DTPicker DTFechaHasta 
      Height          =   375
      Left            =   4320
      TabIndex        =   1
      Top             =   240
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CalendarBackColor=   -2147483639
      CalendarForeColor=   -2147483625
      CalendarTitleBackColor=   8388608
      CalendarTitleForeColor=   -2147483639
      CalendarTrailingForeColor=   -2147483626
      Format          =   50003969
      CurrentDate     =   36892
   End
   Begin MSComctlLib.ProgressBar BarraProgreso 
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1200
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
      Min             =   1e-4
      Scrolling       =   1
   End
   Begin VB.Label LComentario 
      Caption         =   "Fecha Desde"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   7
      Top             =   600
      Width           =   5745
   End
   Begin VB.Label LFechaHasta 
      Caption         =   "Hasta"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3360
      TabIndex        =   5
      Top             =   240
      Width           =   945
   End
   Begin VB.Label LFechaDesde 
      Caption         =   "Fecha Desde"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   240
      Width           =   1425
   End
   Begin VB.Label LProgreso 
      Caption         =   "Fecha Desde"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   5745
   End
End
Attribute VB_Name = "FrmFiltroImpCotMercados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BAceptar_Click()

VarFechaDesde = DTFechaDesde.Value
VarFechaHasta = DTFechaHasta.Value

LProgreso.Visible = True
LComentario.Visible = True
BarraProgreso.Visible = True

LFechaDesde.Visible = False
LFechaHasta.Visible = False
DTFechaDesde.Visible = False
DTFechaHasta.Visible = False
BAceptar.Visible = False
BCancelar.Visible = False

BarraProgreso.Min = 0
BarraProgreso.Max = 100
BarraProgreso.Value = 0

FrmPrincipal.StatusBar1.Panels(3).Visible = True
      
LProgreso.Caption = "1 de 3 - Cargando Mercados para importacion"
FrmPrincipal.StatusBar1.Panels(3).Text = "1 de 3 - Cargando Mercados para importacion"

BarraProgreso.Value = 0

LComentario.Caption = ""

DoEvents

CargarMercadosImportacionCotizaciones ("Yahoo")

LProgreso.Caption = "2 de 3 - Descargando Mercados para importacion"
FrmPrincipal.StatusBar1.Panels(3).Text = "2 de 3 - Descargando Mercados para importacion"

BarraProgreso.Value = 5

LComentario.Caption = ""

DoEvents

DescargarMercadosImportacionCotizaciones ("Yahoo")

BarraProgreso.Value = 20

LProgreso.Caption = "3 de 3 - Recogiendo Mercados para importacion"
FrmPrincipal.StatusBar1.Panels(3).Text = "3 de 3 - Recogiendo Mercados para importacion"


LComentario.Caption = ""

DoEvents

RecogerMercadosImportacionCotizacionesFSO

BarraProgreso.Value = 100

DoEvents

FrmPrincipal.StatusBar1.Panels(3).Visible = False

Unload Me

End Sub

Private Sub BCancelar_Click()
Unload Me
End Sub

Private Sub Form_Load()

DTFechaDesde.Value = Now()
DTFechaHasta.Value = Now()

LProgreso.Visible = False
LComentario.Visible = False
BarraProgreso.Visible = False

LFechaDesde.Visible = True
LFechaHasta.Visible = True
DTFechaDesde.Visible = True
DTFechaHasta.Visible = True
BAceptar.Visible = True
BCancelar.Visible = True

End Sub

