VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.MDIForm FrmPrincipal 
   BackColor       =   &H80000013&
   Caption         =   "MDIForm1"
   ClientHeight    =   6255
   ClientLeft      =   165
   ClientTop       =   765
   ClientWidth     =   10965
   Icon            =   "FrmPrincipal.frx":0000
   LinkTopic       =   "MDIForm1"
   MousePointer    =   2  'Cross
   StartUpPosition =   3  'Windows Default
   WhatsThisHelp   =   -1  'True
   WindowState     =   2  'Maximized
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   5760
      Width           =   10965
      _ExtentX        =   19341
      _ExtentY        =   873
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   1058
            MinWidth        =   1058
            TextSave        =   "18:12"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "28/09/2016"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            AutoSize        =   1
            Object.Visible         =   0   'False
            Object.Width           =   15928
            Key             =   "sbrText"
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
   Begin VB.Menu MenuArchivo 
      Caption         =   "Archivo"
      Index           =   0
      Begin VB.Menu ConexionSQL 
         Caption         =   "Parametrización conexión SQL"
         Enabled         =   0   'False
      End
      Begin VB.Menu BackupSQL 
         Caption         =   "Backup SQL"
         Enabled         =   0   'False
      End
      Begin VB.Menu RestaurarSQL 
         Caption         =   "Restaurar SQL"
         Enabled         =   0   'False
      End
      Begin VB.Menu Separador 
         Caption         =   "-"
      End
      Begin VB.Menu Salir 
         Caption         =   "Salir"
         Index           =   1
      End
   End
   Begin VB.Menu General 
      Caption         =   "General"
      Begin VB.Menu CfgIndicadores 
         Caption         =   "Configuración Indicadores"
      End
      Begin VB.Menu LinGeneral 
         Caption         =   "-"
      End
      Begin VB.Menu cfgExportCSV 
         Caption         =   "Exportación ficheros CSV para Viewer"
      End
   End
   Begin VB.Menu MenuMercados 
      Caption         =   "Mercados"
      Index           =   0
      Begin VB.Menu MercadosSegmentados 
         Caption         =   "Mercados segmentados"
         Shortcut        =   {F1}
      End
      Begin VB.Menu Lin2Mercados 
         Caption         =   "-"
      End
      Begin VB.Menu DescargarCotHisMercados 
         Caption         =   "Descargar cotizaciones históricas"
         Shortcut        =   {F2}
      End
      Begin VB.Menu DescargarMercados 
         Caption         =   "Descargar cotizaciones actuales"
         Enabled         =   0   'False
         Shortcut        =   {F3}
      End
      Begin VB.Menu Lin3Mercados 
         Caption         =   "-"
      End
      Begin VB.Menu ATecnicoMercados 
         Caption         =   "Realizar análisis técnico"
         Shortcut        =   {F4}
      End
   End
   Begin VB.Menu MenuAcciones 
      Caption         =   "Acciones"
      Index           =   1
      Begin VB.Menu AccionesSegmentadas 
         Caption         =   "Acciones segmentadas"
         Shortcut        =   {F5}
      End
      Begin VB.Menu Lin1Acciones 
         Caption         =   "-"
      End
      Begin VB.Menu CrearAcciones 
         Caption         =   "Crear Acciones desde Mercados"
         Enabled         =   0   'False
      End
      Begin VB.Menu Lin2Acciones 
         Caption         =   "-"
      End
      Begin VB.Menu DescargarCotHis 
         Caption         =   "Descargar cotizaciones históricas"
         Index           =   12
         Shortcut        =   {F6}
      End
      Begin VB.Menu DescargarAcciones 
         Caption         =   "Descargar cotizaciones actuales por Mercado"
         Enabled         =   0   'False
         Index           =   9
         Shortcut        =   {F7}
      End
      Begin VB.Menu DescargarAcciones2 
         Caption         =   "Descargar cotizaciones actuales por Acción"
         Enabled         =   0   'False
         Shortcut        =   ^{F7}
      End
      Begin VB.Menu Lin3Acciones 
         Caption         =   "-"
      End
      Begin VB.Menu ATecnicoAcciones 
         Caption         =   "Realizar análisis técnico"
         Shortcut        =   {F8}
      End
   End
   Begin VB.Menu MenuVentanas 
      Caption         =   "Ventanas"
      Index           =   10
      Begin VB.Menu Cascada 
         Caption         =   "Cascada"
      End
      Begin VB.Menu MVertical 
         Caption         =   "Mosaico vertical"
      End
      Begin VB.Menu MHorizontal 
         Caption         =   "Mosaico horizontal"
      End
      Begin VB.Menu MaxTodas 
         Caption         =   "Maximizar todas"
      End
      Begin VB.Menu MinTodas 
         Caption         =   "Minimizar todas"
      End
      Begin VB.Menu CerrarTodas 
         Caption         =   "Cerrar todas"
      End
   End
   Begin VB.Menu Ayuda 
      Caption         =   "Ayuda"
      Begin VB.Menu AcercaDe 
         Caption         =   "Acerca de SORINVER"
      End
   End
End
Attribute VB_Name = "FrmPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Formularios As Form


Private Sub AccionesSegmentadas_Click()

' Lanzamos el form de muestra de acciones segmentados
Load frmArbolAcciones

End Sub

Private Sub AcercaDe_Click()

' Llamamos al formulario frmAcercaDe
Load frmAcercaDe

End Sub


Private Sub ATecnicoAcciones_Click()

' Llamamos al procedimiento encargado de realizar el análisis técnico a las acciones
AnalisisTecnicoAcciones

End Sub

Private Sub ATecnicoMercados_Click()

' Llamamos al procedimiento encargado de realizar el análisis técnico a los índices
AnalisisTecnicoMercados

End Sub


Private Sub Cascada_Click()

' Modo cascada
Me.Arrange vbCascade

End Sub

Private Sub CerrarTodas_Click()

For Each Formularios In Forms

'Si no es el MDI lo descarga
If Not Formularios Is Me Then

Unload Formularios
End If

Next

End Sub



Private Sub CfgIndicadores_Click()

' Lanzamos el form de parametrización indicadores
Load frmConfiguracionIndicadores

End Sub

Private Sub DescargarCotHis_Click(Index As Integer)

' Lanzamos el form de petición de rango de fechas a descargar
Load FrmFiltroImpCotAcciones

End Sub

Private Sub DescargarCotHisMercados_Click()

' Lanzamos el form de petición de rango de fechas a descargar
Load FrmFiltroImpCotMercados

End Sub

Private Sub MaxTodas_Click()

'Recorre todos los Formularios en un For-Each
For Each Formularios In Forms

'si no es el MDI entonces reestablece la ventana
If Not (Formularios Is Me) Then
Formularios.WindowState = vbNormal
End If

Next

End Sub

Private Sub MDIForm_Load()

'Ponemos el titulo de la aplicación y la versión en la barra superior
Me.Caption = App.Title & " - Versión " & App.Major & "." & App.Minor & "." & App.Revision

'Inicializamos datos de la aplicacion
InicializarAplicacion

'Cargamos contenido de mercados a la matriz
CargarMercados

End Sub


Private Sub MercadosSegmentados_Click()

' Lanzamos el form de muestra de mercados segmentados
Load frmMercadosSegmentados

End Sub



Private Sub MHorizontal_Click()

' Mosaico Horizontal
Me.Arrange vbTileHorizontal

End Sub

Private Sub MinTodas_Click()

'Recorre todos los Formularios en un For-Each
For Each Formularios In Forms

'si no es el MDI entonces lo minimiza
If Not Formularios Is Me Then
Formularios.WindowState = vbMinimized
End If

Next

End Sub

Private Sub MVertical_Click()

' Mosaico Vertical
Me.Arrange vbTileVertical

End Sub


Private Sub Salir_Click(Index As Integer)

' Cerramos el programa
Unload Me

End Sub


