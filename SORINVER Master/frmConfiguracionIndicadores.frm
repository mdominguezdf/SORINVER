VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmConfiguracionIndicadores 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Configuración Indicadores"
   ClientHeight    =   4785
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6705
   Icon            =   "frmConfiguracionIndicadores.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4785
   ScaleWidth      =   6705
   Begin VB.CommandButton BAceptar 
      Caption         =   "Aceptar"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5040
      TabIndex        =   35
      Top             =   4200
      Width           =   1575
   End
   Begin VB.CommandButton BCancelar 
      Caption         =   "Cancelar"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   34
      Top             =   4200
      Width           =   1575
   End
   Begin VB.Frame FraRSIAvisos 
      Caption         =   "Avisos"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   1215
      Left            =   240
      TabIndex        =   30
      Top             =   2400
      Width           =   6255
      Begin VB.CheckBox ChRSI_AvisoSalidaZona 
         Caption         =   "Estrategia Salida Zona"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   33
         Top             =   360
         Width           =   2775
      End
      Begin VB.CheckBox ChRSI_AvisoDivergencia 
         Caption         =   "Estrategia Divergencia"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   32
         Top             =   720
         Width           =   2775
      End
      Begin VB.CheckBox ChRSI_AvisoFailureSwing 
         Caption         =   "Estrategia Failure Swing"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3120
         TabIndex        =   31
         Top             =   360
         Width           =   2775
      End
   End
   Begin VB.Frame FraRSIParametros 
      Caption         =   "Parametros"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   1215
      Left            =   240
      TabIndex        =   23
      Top             =   600
      Width           =   6255
      Begin VB.TextBox TBRSI_Periodo 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   960
         TabIndex        =   26
         Top             =   480
         Width           =   375
      End
      Begin VB.TextBox TBRSI_SobreCompra 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   4080
         TabIndex        =   25
         Top             =   480
         Width           =   375
      End
      Begin VB.TextBox TBRSI_SobreVenta 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   5640
         TabIndex        =   24
         Top             =   480
         Width           =   375
      End
      Begin VB.Label Label14 
         Caption         =   "Periodo"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   29
         Top             =   480
         Width           =   615
      End
      Begin VB.Label Label9 
         Caption         =   "Sobre Compra"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3000
         TabIndex        =   28
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label8 
         Caption         =   "Sobre Venta"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4680
         TabIndex        =   27
         Top             =   480
         Width           =   975
      End
   End
   Begin VB.Frame FraLaneAvisos 
      Caption         =   "Avisos"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   1575
      Left            =   240
      TabIndex        =   2
      Top             =   2400
      Width           =   6255
      Begin VB.CheckBox ChLane_AvisoPopCorn_Rapido 
         Caption         =   "Estrategia Pop Corn en Rapido"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3120
         TabIndex        =   18
         Top             =   1080
         Width           =   2775
      End
      Begin VB.CheckBox ChLane_AvisoSZona_Rapido 
         Caption         =   "Estrategia Salida Zona en Rapido"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3120
         TabIndex        =   17
         Top             =   720
         Width           =   2775
      End
      Begin VB.CheckBox ChLane_AvisoClasica_Rapido 
         Caption         =   "Estrategia Clasica en Rapido"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3120
         TabIndex        =   16
         Top             =   360
         Width           =   2775
      End
      Begin VB.CheckBox ChLane_AvisoPopCorn_Lento 
         Caption         =   "Estrategia Pop Corn en Lento"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   15
         Top             =   1080
         Width           =   2775
      End
      Begin VB.CheckBox ChLane_AvisoSZona_Lento 
         Caption         =   "Estrategia Salida Zona en Lento"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   14
         Top             =   720
         Width           =   2775
      End
      Begin VB.CheckBox ChLane_AvisoClasica_Lento 
         Caption         =   "Estrategia Clasica en Lento"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   13
         Top             =   360
         Width           =   2775
      End
   End
   Begin VB.Frame FraLaneParametros 
      Caption         =   "Parametros"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   1695
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   6255
      Begin VB.TextBox TBLane_SobreVenta 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   2880
         TabIndex        =   21
         Top             =   1080
         Width           =   375
      End
      Begin VB.TextBox TBLane_SobreCompra 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1320
         TabIndex        =   19
         Top             =   1080
         Width           =   375
      End
      Begin VB.TextBox TBLane_DSS 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   5640
         TabIndex        =   11
         Top             =   480
         Width           =   375
      End
      Begin VB.TextBox TBLane_DS 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   4440
         TabIndex        =   9
         Top             =   480
         Width           =   375
      End
      Begin VB.TextBox TBLane_D 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   3120
         TabIndex        =   7
         Top             =   480
         Width           =   375
      End
      Begin VB.TextBox TBLane_K 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   2040
         TabIndex        =   5
         Top             =   480
         Width           =   375
      End
      Begin VB.TextBox TBLane_Periodo 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   960
         TabIndex        =   3
         Top             =   480
         Width           =   375
      End
      Begin VB.Label Label7 
         Caption         =   "Sobre Venta"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1920
         TabIndex        =   22
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label6 
         Caption         =   "Sobre Compra"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   20
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "%DSS"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5160
         TabIndex        =   12
         Top             =   480
         Width           =   495
      End
      Begin VB.Label Label4 
         Caption         =   "%DS"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4080
         TabIndex        =   10
         Top             =   480
         Width           =   375
      End
      Begin VB.Label Label3 
         Caption         =   "%D"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2760
         TabIndex        =   8
         Top             =   480
         Width           =   255
      End
      Begin VB.Label Label2 
         Caption         =   "%K"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1680
         TabIndex        =   6
         Top             =   480
         Width           =   255
      End
      Begin VB.Label Label1 
         Caption         =   "Periodo"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   480
         Width           =   615
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   4095
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   7223
      MultiRow        =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Estocastico"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "RSI"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmConfiguracionIndicadores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BAceptar_Click()

If IsNumeric(TBLane_Periodo.Text) Then Lane_Periodo = TBLane_Periodo.Text
If IsNumeric(TBLane_K.Text) Then Lane_K = TBLane_K.Text
If IsNumeric(TBLane_D.Text) Then Lane_D = TBLane_D.Text
If IsNumeric(TBLane_DS.Text) Then Lane_DS = TBLane_DS.Text
If IsNumeric(TBLane_DSS.Text) Then Lane_DSS = TBLane_DSS.Text
If IsNumeric(TBLane_SobreCompra.Text) Then Lane_SobreCompra = TBLane_SobreCompra.Text
If IsNumeric(TBLane_SobreVenta.Text) Then Lane_SobreVenta = TBLane_SobreVenta.Text

If ChLane_AvisoClasica_Lento.Value = 1 Then Lane_AvisoClasica_Lento = True Else Lane_AvisoClasica_Lento = False
If ChLane_AvisoSZona_Lento.Value = 1 Then Lane_AvisoSZona_Lento = True Else Lane_AvisoSZona_Lento = False
If ChLane_AvisoPopCorn_Lento.Value = 1 Then Lane_AvisoPopCorn_Lento = True Else Lane_AvisoPopCorn_Lento = False
If ChLane_AvisoClasica_Rapido.Value = 1 Then Lane_AvisoClasica_Rapido = True Else Lane_AvisoClasica_Rapido = False
If ChLane_AvisoSZona_Rapido.Value = 1 Then Lane_AvisoSZona_Rapido = True Else Lane_AvisoSZona_Rapido = False
If ChLane_AvisoPopCorn_Rapido.Value = 1 Then Lane_AvisoPopCorn_Rapido = True Else Lane_AvisoPopCorn_Rapido = False


If IsNumeric(TBRSI_Periodo.Text) Then RSI_Periodo = TBRSI_Periodo.Text
If IsNumeric(TBRSI_SobreCompra.Text) Then RSI_SobreCompra = TBRSI_SobreCompra.Text
If IsNumeric(TBRSI_SobreVenta.Text) Then RSI_SobreVenta = TBRSI_SobreVenta.Text

If ChRSI_AvisoSalidaZona.Value = 1 Then RSI_AvisoSalidaZona = True Else RSI_AvisoSalidaZona = False
If ChRSI_AvisoFailureSwing.Value = 1 Then RSI_AvisoFailureSwing = True Else RSI_AvisoFailureSwing = False
If ChRSI_AvisoDivergencia.Value = 1 Then RSI_AvisoDivergencia = True Else RSI_AvisoDivergencia = False

GuardarDefectosSQL

Unload Me

End Sub

Private Sub BCancelar_Click()

Unload Me

End Sub

Private Sub Form_Load()

FraLaneParametros.Visible = True
FraLaneAvisos.Visible = True
FraRSIParametros.Visible = False
FraRSIAvisos.Visible = False

TBLane_Periodo.Text = Lane_Periodo
TBLane_K.Text = Lane_K
TBLane_D.Text = Lane_D
TBLane_DS.Text = Lane_DS
TBLane_DSS.Text = Lane_DSS
TBLane_SobreCompra.Text = Lane_SobreCompra
TBLane_SobreVenta.Text = Lane_SobreVenta

If Lane_AvisoClasica_Lento Then ChLane_AvisoClasica_Lento.Value = 1 Else ChLane_AvisoClasica_Lento.Value = 0
If Lane_AvisoSZona_Lento Then ChLane_AvisoSZona_Lento.Value = 1 Else ChLane_AvisoSZona_Lento.Value = 0
If Lane_AvisoPopCorn_Lento Then ChLane_AvisoPopCorn_Lento.Value = 1 Else ChLane_AvisoPopCorn_Lento.Value = 0
If Lane_AvisoClasica_Rapido Then ChLane_AvisoClasica_Rapido.Value = 1 Else ChLane_AvisoClasica_Rapido.Value = 0
If Lane_AvisoSZona_Rapido Then ChLane_AvisoSZona_Rapido.Value = 1 Else ChLane_AvisoSZona_Rapido.Value = 0
If Lane_AvisoPopCorn_Rapido Then ChLane_AvisoPopCorn_Rapido.Value = 1 Else ChLane_AvisoPopCorn_Rapido.Value = 0

TBRSI_Periodo.Text = RSI_Periodo
TBRSI_SobreCompra.Text = RSI_SobreCompra
TBRSI_SobreVenta.Text = RSI_SobreVenta

If RSI_AvisoSalidaZona Then ChRSI_AvisoSalidaZona.Value = 1 Else ChRSI_AvisoSalidaZona.Value = 0
If RSI_AvisoFailureSwing Then ChRSI_AvisoFailureSwing.Value = 1 Else ChRSI_AvisoFailureSwing.Value = 0
If RSI_AvisoDivergencia Then ChRSI_AvisoDivergencia.Value = 1 Else ChRSI_AvisoDivergencia.Value = 0

End Sub

Private Sub TabStrip1_Click()


If TabStrip1.SelectedItem = "RSI" Then

   FraLaneParametros.Visible = False
   FraLaneAvisos.Visible = False
   FraRSIParametros.Visible = True
   FraRSIAvisos.Visible = True
   
Else
   
   FraLaneParametros.Visible = True
   FraLaneAvisos.Visible = True
   FraRSIParametros.Visible = False
   FraRSIAvisos.Visible = False
   
End If

End Sub

