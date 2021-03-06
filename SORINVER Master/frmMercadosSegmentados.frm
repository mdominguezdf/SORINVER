VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmMercadosSegmentados 
   Caption         =   "Mercados Segmentados"
   ClientHeight    =   6705
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7650
   Icon            =   "frmMercadosSegmentados.frx":0000
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   6705
   ScaleWidth      =   7650
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   960
      Top             =   960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMercadosSegmentados.frx":0442
            Key             =   "Accion"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMercadosSegmentados.frx":0894
            Key             =   "Cerrado"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMercadosSegmentados.frx":0CE6
            Key             =   "Abierto"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMercadosSegmentados.frx":1138
            Key             =   "LogoSM"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMercadosSegmentados.frx":158A
            Key             =   "Zona"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView VistaArbol 
      Height          =   7095
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   12515
      _Version        =   393217
      Indentation     =   353
      Style           =   7
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmMercadosSegmentados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()

CargarArbolMercados

End Sub

Private Sub Form_Resize()

VistaArbol.Width = frmMercadosSegmentados.Width - 100
VistaArbol.Height = frmMercadosSegmentados.Height - 500

VistaArbol.Refresh

End Sub

Private Sub VistaArbol_NodeClick(ByVal Node As MSComctlLib.Node)

   If Node.Children = 0 Then
   
      If ControlDblClick = VistaArbol.SelectedItem.Key Then

         FichaValor "Mercados", VistaArbol.SelectedItem.Text, Right(Node.Key, Len(Node.Key) - InStr(1, Node.Key, ".", vbTextCompare))
                  
      Else
      
         ControlDblClick = VistaArbol.SelectedItem.Key
      
      End If
      
   End If

End Sub
