VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmFichaValor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "frmFichaValor"
   ClientHeight    =   4845
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8055
   BeginProperty Font 
      Name            =   "Trebuchet MS"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmFichaValor.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4845
   ScaleWidth      =   8055
   Begin VB.CommandButton BExportarWeka 
      Caption         =   "Exportar Weka"
      Height          =   495
      Left            =   120
      TabIndex        =   190
      Top             =   4320
      Width           =   1215
   End
   Begin VB.Frame fraTendenciasBajistas 
      Caption         =   "Bajistas"
      Enabled         =   0   'False
      ForeColor       =   &H008080FF&
      Height          =   1695
      Left            =   240
      TabIndex        =   158
      Top             =   2280
      Width           =   7575
      Begin VB.TextBox TBTBL 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
         Height          =   225
         Index           =   1
         Left            =   720
         TabIndex        =   179
         Top             =   1320
         Width           =   975
      End
      Begin VB.TextBox TBTBL 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
         ForeColor       =   &H008080FF&
         Height          =   225
         Index           =   2
         Left            =   3960
         TabIndex        =   178
         Top             =   1320
         Width           =   975
      End
      Begin VB.TextBox TBTBL 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
         Height          =   225
         Index           =   3
         Left            =   2880
         TabIndex        =   177
         Top             =   1320
         Width           =   975
      End
      Begin VB.TextBox TBTBL 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
         ForeColor       =   &H008080FF&
         Height          =   225
         Index           =   0
         Left            =   1800
         TabIndex        =   176
         Top             =   1320
         Width           =   975
      End
      Begin VB.TextBox TBTBL 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
         Height          =   225
         Index           =   4
         Left            =   5040
         TabIndex        =   175
         Top             =   1320
         Width           =   855
      End
      Begin VB.TextBox TBTBL 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
         Height          =   225
         Index           =   5
         Left            =   6000
         TabIndex        =   174
         Top             =   1320
         Width           =   735
      End
      Begin VB.TextBox TBTBL 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
         Height          =   225
         Index           =   6
         Left            =   6840
         TabIndex        =   173
         Top             =   1320
         Width           =   615
      End
      Begin VB.TextBox TBTBM 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
         Height          =   225
         Index           =   6
         Left            =   6840
         TabIndex        =   172
         Top             =   960
         Width           =   615
      End
      Begin VB.TextBox TBTBM 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
         Height          =   225
         Index           =   5
         Left            =   6000
         TabIndex        =   171
         Top             =   960
         Width           =   735
      End
      Begin VB.TextBox TBTBM 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
         Height          =   225
         Index           =   4
         Left            =   5040
         TabIndex        =   170
         Top             =   960
         Width           =   855
      End
      Begin VB.TextBox TBTBM 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
         ForeColor       =   &H008080FF&
         Height          =   225
         Index           =   0
         Left            =   1800
         TabIndex        =   169
         Top             =   960
         Width           =   975
      End
      Begin VB.TextBox TBTBM 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
         Height          =   225
         Index           =   3
         Left            =   2880
         TabIndex        =   168
         Top             =   960
         Width           =   975
      End
      Begin VB.TextBox TBTBM 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
         ForeColor       =   &H008080FF&
         Height          =   225
         Index           =   2
         Left            =   3960
         TabIndex        =   167
         Top             =   960
         Width           =   975
      End
      Begin VB.TextBox TBTBM 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
         Height          =   225
         Index           =   1
         Left            =   720
         TabIndex        =   166
         Top             =   960
         Width           =   975
      End
      Begin VB.TextBox TBTBC 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
         Height          =   225
         Index           =   6
         Left            =   6840
         TabIndex        =   165
         Top             =   600
         Width           =   615
      End
      Begin VB.TextBox TBTBC 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
         Height          =   225
         Index           =   5
         Left            =   6000
         TabIndex        =   164
         Top             =   600
         Width           =   735
      End
      Begin VB.TextBox TBTBC 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
         Height          =   225
         Index           =   4
         Left            =   5040
         TabIndex        =   163
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox TBTBC 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
         ForeColor       =   &H008080FF&
         Height          =   225
         Index           =   0
         Left            =   1800
         TabIndex        =   162
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox TBTBC 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
         Height          =   225
         Index           =   3
         Left            =   2880
         TabIndex        =   161
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox TBTBC 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
         ForeColor       =   &H008080FF&
         Height          =   225
         Index           =   2
         Left            =   3960
         TabIndex        =   160
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox TBTBC 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
         Height          =   225
         Index           =   1
         Left            =   720
         TabIndex        =   159
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label56 
         Alignment       =   1  'Right Justify
         Caption         =   "Comienzo"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   720
         TabIndex        =   187
         Top             =   240
         Width           =   975
      End
      Begin VB.Label LTBTBC 
         Alignment       =   1  'Right Justify
         Caption         =   "Corto"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   186
         Top             =   600
         Width           =   495
      End
      Begin VB.Label LTBTBM 
         Alignment       =   1  'Right Justify
         Caption         =   "Medio"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   185
         Top             =   960
         Width           =   495
      End
      Begin VB.Label LTBTBL 
         Alignment       =   1  'Right Justify
         Caption         =   "Largo"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   184
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label Label52 
         Alignment       =   1  'Right Justify
         Caption         =   "Ultimo"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2880
         TabIndex        =   183
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label51 
         Alignment       =   1  'Right Justify
         Caption         =   "% Pendiente"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4680
         TabIndex        =   182
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label50 
         Alignment       =   1  'Right Justify
         Caption         =   "% Acum."
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6000
         TabIndex        =   181
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label49 
         Alignment       =   1  'Right Justify
         Caption         =   "Días"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6840
         TabIndex        =   180
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Frame fraTendenciasAlcistas 
      Caption         =   "Alcistas"
      Enabled         =   0   'False
      ForeColor       =   &H8000000D&
      Height          =   1695
      Left            =   240
      TabIndex        =   128
      Top             =   600
      Width           =   7575
      Begin VB.TextBox TBTAC 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
         Height          =   225
         Index           =   1
         Left            =   720
         TabIndex        =   153
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox TBTAC 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   225
         Index           =   2
         Left            =   3960
         TabIndex        =   152
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox TBTAC 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
         Height          =   225
         Index           =   3
         Left            =   2880
         TabIndex        =   151
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox TBTAC 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   225
         Index           =   0
         Left            =   1800
         TabIndex        =   150
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox TBTAC 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
         Height          =   225
         Index           =   4
         Left            =   5040
         TabIndex        =   149
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox TBTAC 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
         Height          =   225
         Index           =   5
         Left            =   6000
         TabIndex        =   148
         Top             =   600
         Width           =   735
      End
      Begin VB.TextBox TBTAC 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
         Height          =   225
         Index           =   6
         Left            =   6840
         TabIndex        =   147
         Top             =   600
         Width           =   615
      End
      Begin VB.TextBox TBTAM 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
         Height          =   225
         Index           =   1
         Left            =   720
         TabIndex        =   146
         Top             =   960
         Width           =   975
      End
      Begin VB.TextBox TBTAM 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   225
         Index           =   2
         Left            =   3960
         TabIndex        =   145
         Top             =   960
         Width           =   975
      End
      Begin VB.TextBox TBTAM 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
         Height          =   225
         Index           =   3
         Left            =   2880
         TabIndex        =   144
         Top             =   960
         Width           =   975
      End
      Begin VB.TextBox TBTAM 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   225
         Index           =   0
         Left            =   1800
         TabIndex        =   143
         Top             =   960
         Width           =   975
      End
      Begin VB.TextBox TBTAM 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
         Height          =   225
         Index           =   4
         Left            =   5040
         TabIndex        =   142
         Top             =   960
         Width           =   855
      End
      Begin VB.TextBox TBTAM 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
         Height          =   225
         Index           =   5
         Left            =   6000
         TabIndex        =   141
         Top             =   960
         Width           =   735
      End
      Begin VB.TextBox TBTAM 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
         Height          =   225
         Index           =   6
         Left            =   6840
         TabIndex        =   140
         Top             =   960
         Width           =   615
      End
      Begin VB.TextBox TBTAL 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
         Height          =   225
         Index           =   6
         Left            =   6840
         TabIndex        =   139
         Top             =   1320
         Width           =   615
      End
      Begin VB.TextBox TBTAL 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
         Height          =   225
         Index           =   5
         Left            =   6000
         TabIndex        =   138
         Top             =   1320
         Width           =   735
      End
      Begin VB.TextBox TBTAL 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
         Height          =   225
         Index           =   4
         Left            =   5040
         TabIndex        =   137
         Top             =   1320
         Width           =   855
      End
      Begin VB.TextBox TBTAL 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   225
         Index           =   0
         Left            =   1800
         TabIndex        =   132
         Top             =   1320
         Width           =   975
      End
      Begin VB.TextBox TBTAL 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
         Height          =   225
         Index           =   3
         Left            =   2880
         TabIndex        =   131
         Top             =   1320
         Width           =   975
      End
      Begin VB.TextBox TBTAL 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   225
         Index           =   2
         Left            =   3960
         TabIndex        =   130
         Top             =   1320
         Width           =   975
      End
      Begin VB.TextBox TBTAL 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
         Height          =   225
         Index           =   1
         Left            =   720
         TabIndex        =   129
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label Label48 
         Alignment       =   1  'Right Justify
         Caption         =   "Días"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6840
         TabIndex        =   157
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label47 
         Alignment       =   1  'Right Justify
         Caption         =   "% Acum."
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6000
         TabIndex        =   156
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label46 
         Alignment       =   1  'Right Justify
         Caption         =   "% Pendiente"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4680
         TabIndex        =   155
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label44 
         Alignment       =   1  'Right Justify
         Caption         =   "Ultimo"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2880
         TabIndex        =   154
         Top             =   240
         Width           =   975
      End
      Begin VB.Label LTBTAL 
         Alignment       =   1  'Right Justify
         Caption         =   "Largo"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   136
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label LTBTAM 
         Alignment       =   1  'Right Justify
         Caption         =   "Medio"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   135
         Top             =   960
         Width           =   495
      End
      Begin VB.Label LTBTAC 
         Alignment       =   1  'Right Justify
         Caption         =   "Corto"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   134
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label43 
         Alignment       =   1  'Right Justify
         Caption         =   "Comienzo"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   720
         TabIndex        =   133
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame fraGeneralControles 
      Caption         =   "Controles cotizaciones"
      ForeColor       =   &H8000000D&
      Height          =   1215
      Left            =   240
      TabIndex        =   1
      Top             =   2760
      Width           =   6255
      Begin VB.CheckBox ChControlValorCot 
         Caption         =   "Control cot. actuales acciones"
         Height          =   375
         Left            =   3120
         TabIndex        =   5
         Top             =   720
         Width           =   2775
      End
      Begin VB.CheckBox ChControlCot 
         Caption         =   "Control cot. actuales mercados"
         Height          =   375
         Left            =   3120
         TabIndex        =   4
         Top             =   360
         Width           =   2775
      End
      Begin VB.CheckBox ChControlValorHis 
         Caption         =   "Imp. cot. historicas acciones"
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   720
         Width           =   2775
      End
      Begin VB.CheckBox ChControlHis 
         Caption         =   "Imp. cot. historicas mercados"
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   2775
      End
   End
   Begin VB.Frame fraSoportesyResistencias 
      Enabled         =   0   'False
      ForeColor       =   &H8000000D&
      Height          =   2415
      Left            =   240
      TabIndex        =   73
      Top             =   720
      Width           =   6255
      Begin VB.TextBox TBSop200 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
         Height          =   225
         Index           =   2
         Left            =   4800
         TabIndex        =   90
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox TBSop200 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   225
         Index           =   1
         Left            =   4800
         TabIndex        =   89
         Top             =   1800
         Width           =   1095
      End
      Begin VB.TextBox TBSop200 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
         Height          =   225
         Index           =   0
         Left            =   4800
         TabIndex        =   88
         Top             =   1440
         Width           =   1095
      End
      Begin VB.TextBox TBSop50 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
         Height          =   225
         Index           =   2
         Left            =   3120
         TabIndex        =   87
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox TBSop50 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   225
         Index           =   1
         Left            =   3120
         TabIndex        =   86
         Top             =   1800
         Width           =   1095
      End
      Begin VB.TextBox TBSop50 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
         Height          =   225
         Index           =   0
         Left            =   3120
         TabIndex        =   85
         Top             =   1440
         Width           =   1095
      End
      Begin VB.TextBox TBSop20 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
         Height          =   225
         Index           =   2
         Left            =   1440
         TabIndex        =   84
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox TBSop20 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   225
         Index           =   1
         Left            =   1440
         TabIndex        =   83
         Top             =   1800
         Width           =   1095
      End
      Begin VB.TextBox TBSop20 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
         Height          =   225
         Index           =   0
         Left            =   1440
         TabIndex        =   82
         Top             =   1440
         Width           =   1095
      End
      Begin VB.TextBox TBSop20 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   225
         Index           =   3
         Left            =   1440
         TabIndex        =   81
         Top             =   1080
         Width           =   1095
      End
      Begin VB.TextBox TBSop50 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   225
         Index           =   3
         Left            =   3120
         TabIndex        =   80
         Top             =   1080
         Width           =   1095
      End
      Begin VB.TextBox TBSop200 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   225
         Index           =   3
         Left            =   4800
         TabIndex        =   74
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label30 
         Alignment       =   1  'Right Justify
         Caption         =   "200"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5280
         TabIndex        =   79
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label29 
         Alignment       =   1  'Right Justify
         Caption         =   "50"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3600
         TabIndex        =   78
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label28 
         Alignment       =   1  'Right Justify
         Caption         =   "20"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1920
         TabIndex        =   77
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label26 
         Alignment       =   1  'Right Justify
         Caption         =   "Soporte"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   76
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label Label25 
         Alignment       =   1  'Right Justify
         Caption         =   "Resistencia"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   75
         Top             =   720
         Width           =   1095
      End
   End
   Begin VB.Frame fraVelas 
      Enabled         =   0   'False
      ForeColor       =   &H8000000D&
      Height          =   2055
      Left            =   240
      TabIndex        =   57
      Top             =   720
      Width           =   6255
      Begin VB.TextBox TBVela200 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
         Height          =   225
         Index           =   1
         Left            =   4440
         TabIndex        =   72
         Top             =   1440
         Width           =   1455
      End
      Begin VB.TextBox TBVela200 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
         Height          =   225
         Index           =   0
         Left            =   4440
         TabIndex        =   71
         Top             =   1080
         Width           =   1455
      End
      Begin VB.TextBox TBVela50 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
         Height          =   225
         Index           =   1
         Left            =   2760
         TabIndex        =   70
         Top             =   1440
         Width           =   1455
      End
      Begin VB.TextBox TBVela50 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
         Height          =   225
         Index           =   0
         Left            =   2760
         TabIndex        =   69
         Top             =   1080
         Width           =   1455
      End
      Begin VB.TextBox TBVela20 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
         Height          =   225
         Index           =   1
         Left            =   1080
         TabIndex        =   68
         Top             =   1440
         Width           =   1455
      End
      Begin VB.TextBox TBVela20 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
         Height          =   225
         Index           =   0
         Left            =   1080
         TabIndex        =   67
         Top             =   1080
         Width           =   1455
      End
      Begin VB.TextBox TBVela50 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
         Height          =   225
         Index           =   2
         Left            =   2760
         TabIndex        =   60
         Top             =   720
         Width           =   1455
      End
      Begin VB.TextBox TBVela20 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
         Height          =   225
         Index           =   2
         Left            =   1080
         TabIndex        =   59
         Top             =   720
         Width           =   1455
      End
      Begin VB.TextBox TBVela200 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
         Height          =   225
         Index           =   2
         Left            =   4440
         TabIndex        =   58
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label24 
         Alignment       =   1  'Right Justify
         Caption         =   "Máxima"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   66
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label23 
         Alignment       =   1  'Right Justify
         Caption         =   "Media"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   65
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label Label22 
         Alignment       =   1  'Right Justify
         Caption         =   "Mínima"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   64
         Top             =   1440
         Width           =   615
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         Caption         =   "20"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1920
         TabIndex        =   63
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         Caption         =   "50"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3600
         TabIndex        =   62
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         Caption         =   "200"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5280
         TabIndex        =   61
         Top             =   360
         Width           =   495
      End
   End
   Begin VB.Frame fraVolumen 
      Enabled         =   0   'False
      ForeColor       =   &H8000000D&
      Height          =   2055
      Left            =   240
      TabIndex        =   41
      Top             =   720
      Width           =   6255
      Begin VB.TextBox TBVol200 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
         Height          =   225
         Index           =   2
         Left            =   4440
         TabIndex        =   56
         Top             =   720
         Width           =   1455
      End
      Begin VB.TextBox TBVol200 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
         Height          =   225
         Index           =   1
         Left            =   4440
         TabIndex        =   55
         Top             =   1440
         Width           =   1455
      End
      Begin VB.TextBox TBVol50 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
         Height          =   225
         Index           =   0
         Left            =   2760
         TabIndex        =   54
         Top             =   1080
         Width           =   1455
      End
      Begin VB.TextBox TBVol20 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
         Height          =   225
         Index           =   0
         Left            =   1080
         TabIndex        =   47
         Top             =   1080
         Width           =   1455
      End
      Begin VB.TextBox TBVol200 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
         Height          =   225
         Index           =   0
         Left            =   4440
         TabIndex        =   46
         Top             =   1080
         Width           =   1455
      End
      Begin VB.TextBox TBVol20 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
         Height          =   225
         Index           =   1
         Left            =   1080
         TabIndex        =   45
         Top             =   1440
         Width           =   1455
      End
      Begin VB.TextBox TBVol20 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
         Height          =   225
         Index           =   2
         Left            =   1080
         TabIndex        =   44
         Top             =   720
         Width           =   1455
      End
      Begin VB.TextBox TBVol50 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
         Height          =   225
         Index           =   1
         Left            =   2760
         TabIndex        =   43
         Top             =   1440
         Width           =   1455
      End
      Begin VB.TextBox TBVol50 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
         Height          =   225
         Index           =   2
         Left            =   2760
         TabIndex        =   42
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "200"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5280
         TabIndex        =   53
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         Caption         =   "50"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3600
         TabIndex        =   52
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         Caption         =   "20"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1920
         TabIndex        =   51
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         Caption         =   "Mínimo"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   50
         Top             =   1440
         Width           =   615
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         Caption         =   "Medio"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   49
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         Caption         =   "Máximo"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   48
         Top             =   720
         Width           =   615
      End
   End
   Begin VB.Frame fraIndicadores 
      Enabled         =   0   'False
      ForeColor       =   &H8000000D&
      Height          =   3255
      Left            =   240
      TabIndex        =   91
      Top             =   720
      Width           =   6255
      Begin VB.TextBox TBRSI 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
         Height          =   225
         Index           =   4
         Left            =   5160
         TabIndex        =   124
         Top             =   2640
         Width           =   735
      End
      Begin VB.TextBox TBRSI 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
         Height          =   225
         Index           =   3
         Left            =   5160
         TabIndex        =   123
         Top             =   2280
         Width           =   735
      End
      Begin VB.TextBox TBRSI 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
         Height          =   225
         Index           =   2
         Left            =   5160
         TabIndex        =   122
         Top             =   1920
         Width           =   735
      End
      Begin VB.TextBox TBRSI 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
         Height          =   225
         Index           =   1
         Left            =   5160
         TabIndex        =   121
         Top             =   1560
         Width           =   735
      End
      Begin VB.TextBox TBDSS 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
         Height          =   225
         Index           =   4
         Left            =   3960
         TabIndex        =   120
         Top             =   2640
         Width           =   735
      End
      Begin VB.TextBox TBDSS 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
         Height          =   225
         Index           =   3
         Left            =   3960
         TabIndex        =   119
         Top             =   2280
         Width           =   735
      End
      Begin VB.TextBox TBDSS 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
         Height          =   225
         Index           =   2
         Left            =   3960
         TabIndex        =   118
         Top             =   1920
         Width           =   735
      End
      Begin VB.TextBox TBDSS 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
         Height          =   225
         Index           =   1
         Left            =   3960
         TabIndex        =   117
         Top             =   1560
         Width           =   735
      End
      Begin VB.TextBox TBDS 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
         Height          =   225
         Index           =   4
         Left            =   3000
         TabIndex        =   116
         Top             =   2640
         Width           =   735
      End
      Begin VB.TextBox TBDS 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
         Height          =   225
         Index           =   3
         Left            =   3000
         TabIndex        =   115
         Top             =   2280
         Width           =   735
      End
      Begin VB.TextBox TBDS 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
         Height          =   225
         Index           =   2
         Left            =   3000
         TabIndex        =   114
         Top             =   1920
         Width           =   735
      End
      Begin VB.TextBox TBDS 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
         Height          =   225
         Index           =   1
         Left            =   3000
         TabIndex        =   113
         Top             =   1560
         Width           =   735
      End
      Begin VB.TextBox TBD 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
         Height          =   225
         Index           =   4
         Left            =   2040
         TabIndex        =   112
         Top             =   2640
         Width           =   735
      End
      Begin VB.TextBox TBD 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
         Height          =   225
         Index           =   3
         Left            =   2040
         TabIndex        =   111
         Top             =   2280
         Width           =   735
      End
      Begin VB.TextBox TBD 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
         Height          =   225
         Index           =   2
         Left            =   2040
         TabIndex        =   110
         Top             =   1920
         Width           =   735
      End
      Begin VB.TextBox TBD 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
         Height          =   225
         Index           =   1
         Left            =   2040
         TabIndex        =   109
         Top             =   1560
         Width           =   735
      End
      Begin VB.TextBox TBRSI 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
         Height          =   225
         Index           =   0
         Left            =   5160
         TabIndex        =   108
         Top             =   1200
         Width           =   735
      End
      Begin VB.TextBox TBDSS 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
         Height          =   225
         Index           =   0
         Left            =   3960
         TabIndex        =   107
         Top             =   1200
         Width           =   735
      End
      Begin VB.TextBox TBDS 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
         Height          =   225
         Index           =   0
         Left            =   3000
         TabIndex        =   106
         Top             =   1200
         Width           =   735
      End
      Begin VB.TextBox TBD 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
         Height          =   225
         Index           =   0
         Left            =   2040
         TabIndex        =   105
         Top             =   1200
         Width           =   735
      End
      Begin VB.TextBox TBK 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
         Height          =   225
         Index           =   4
         Left            =   1080
         TabIndex        =   104
         Top             =   2640
         Width           =   735
      End
      Begin VB.TextBox TBK 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
         Height          =   225
         Index           =   3
         Left            =   1080
         TabIndex        =   103
         Top             =   2280
         Width           =   735
      End
      Begin VB.TextBox TBK 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
         Height          =   225
         Index           =   2
         Left            =   1080
         TabIndex        =   102
         Top             =   1920
         Width           =   735
      End
      Begin VB.TextBox TBK 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
         Height          =   225
         Index           =   1
         Left            =   1080
         TabIndex        =   101
         Top             =   1560
         Width           =   735
      End
      Begin VB.TextBox TBK 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
         Height          =   225
         Index           =   0
         Left            =   1080
         TabIndex        =   92
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label Label40 
         Alignment       =   2  'Center
         Caption         =   "Estocastico (Lane)"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1200
         TabIndex        =   127
         Top             =   480
         Width           =   3375
      End
      Begin VB.Label Label39 
         Alignment       =   1  'Right Justify
         Caption         =   "K"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1200
         TabIndex        =   126
         Top             =   840
         Width           =   495
      End
      Begin VB.Label Label37 
         Alignment       =   1  'Right Justify
         Caption         =   "D"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2160
         TabIndex        =   125
         Top             =   840
         Width           =   495
      End
      Begin VB.Label Label38 
         Alignment       =   1  'Right Justify
         Caption         =   "n"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   100
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label Label36 
         Alignment       =   1  'Right Justify
         Caption         =   "n-1"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   99
         Top             =   1560
         Width           =   615
      End
      Begin VB.Label Label35 
         Alignment       =   1  'Right Justify
         Caption         =   "n-2"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   98
         Top             =   1920
         Width           =   615
      End
      Begin VB.Label Label34 
         Alignment       =   1  'Right Justify
         Caption         =   "n-3"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   97
         Top             =   2280
         Width           =   615
      End
      Begin VB.Label Label33 
         Alignment       =   1  'Right Justify
         Caption         =   "n-4"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   96
         Top             =   2640
         Width           =   615
      End
      Begin VB.Label Label32 
         Alignment       =   1  'Right Justify
         Caption         =   "DS"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3120
         TabIndex        =   95
         Top             =   840
         Width           =   495
      End
      Begin VB.Label Label31 
         Alignment       =   1  'Right Justify
         Caption         =   "DSS"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4080
         TabIndex        =   94
         Top             =   840
         Width           =   495
      End
      Begin VB.Label Label27 
         Alignment       =   1  'Right Justify
         Caption         =   "RSI"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5280
         TabIndex        =   93
         Top             =   840
         Width           =   495
      End
   End
   Begin VB.Frame fraMM 
      Enabled         =   0   'False
      ForeColor       =   &H8000000D&
      Height          =   3255
      Left            =   240
      TabIndex        =   13
      Top             =   720
      Width           =   6255
      Begin VB.TextBox TBMM200 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
         Height          =   225
         Index           =   4
         Left            =   4560
         TabIndex        =   40
         Text            =   "ALCISTA"
         Top             =   2640
         Width           =   1335
      End
      Begin VB.TextBox TBMM200 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
         Height          =   225
         Index           =   3
         Left            =   4560
         TabIndex        =   39
         Text            =   "ALCISTA"
         Top             =   2280
         Width           =   1335
      End
      Begin VB.TextBox TBMM200 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
         Height          =   225
         Index           =   2
         Left            =   4560
         TabIndex        =   38
         Text            =   "ALCISTA"
         Top             =   1920
         Width           =   1335
      End
      Begin VB.TextBox TBMM200 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
         Height          =   225
         Index           =   1
         Left            =   4560
         TabIndex        =   37
         Text            =   "ALCISTA"
         Top             =   1560
         Width           =   1335
      End
      Begin VB.TextBox TBMM50 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
         Height          =   225
         Index           =   4
         Left            =   2880
         TabIndex        =   36
         Text            =   "ALCISTA"
         Top             =   2640
         Width           =   1335
      End
      Begin VB.TextBox TBMM50 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
         Height          =   225
         Index           =   3
         Left            =   2880
         TabIndex        =   35
         Text            =   "ALCISTA"
         Top             =   2280
         Width           =   1335
      End
      Begin VB.TextBox TBMM50 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
         Height          =   225
         Index           =   2
         Left            =   2880
         TabIndex        =   34
         Text            =   "ALCISTA"
         Top             =   1920
         Width           =   1335
      End
      Begin VB.TextBox TBMM50 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
         Height          =   225
         Index           =   1
         Left            =   2880
         TabIndex        =   33
         Text            =   "ALCISTA"
         Top             =   1560
         Width           =   1335
      End
      Begin VB.TextBox TBMM200 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
         Height          =   225
         Index           =   0
         Left            =   4560
         TabIndex        =   32
         Text            =   "ALCISTA"
         Top             =   1200
         Width           =   1335
      End
      Begin VB.TextBox TBMM50 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
         Height          =   225
         Index           =   0
         Left            =   2880
         TabIndex        =   31
         Text            =   "ALCISTA"
         Top             =   1200
         Width           =   1335
      End
      Begin VB.TextBox TBMM20 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
         Height          =   225
         Index           =   4
         Left            =   1200
         TabIndex        =   30
         Text            =   "ALCISTA"
         Top             =   2640
         Width           =   1335
      End
      Begin VB.TextBox TBMM20 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
         Height          =   225
         Index           =   3
         Left            =   1200
         TabIndex        =   29
         Text            =   "ALCISTA"
         Top             =   2280
         Width           =   1335
      End
      Begin VB.TextBox TBMM20 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
         Height          =   225
         Index           =   2
         Left            =   1200
         TabIndex        =   28
         Text            =   "ALCISTA"
         Top             =   1920
         Width           =   1335
      End
      Begin VB.TextBox TBMM20 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
         Height          =   225
         Index           =   1
         Left            =   1200
         TabIndex        =   27
         Text            =   "ALCISTA"
         Top             =   1560
         Width           =   1335
      End
      Begin VB.TextBox TBMM20 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
         Height          =   225
         Index           =   0
         Left            =   1200
         TabIndex        =   26
         Text            =   "ALCISTA"
         Top             =   1200
         Width           =   1335
      End
      Begin VB.TextBox TBSignoMM200 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
         BorderStyle     =   0  'None
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
         Height          =   225
         Left            =   4560
         TabIndex        =   25
         Text            =   "BAJISTA"
         Top             =   720
         Width           =   1335
      End
      Begin VB.TextBox TBSignoMM50 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H0080FF80&
         BorderStyle     =   0  'None
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
         Height          =   225
         Left            =   2880
         TabIndex        =   24
         Text            =   "ALCISTA"
         Top             =   720
         Width           =   1335
      End
      Begin VB.TextBox TBSignoMM20 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H0080FF80&
         BorderStyle     =   0  'None
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
         Height          =   225
         Left            =   1200
         TabIndex        =   23
         Text            =   "ALCISTA"
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         Caption         =   "200"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5280
         TabIndex        =   22
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         Caption         =   "50"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3600
         TabIndex        =   21
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         Caption         =   "20"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1920
         TabIndex        =   20
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "n-4"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   19
         Top             =   2640
         Width           =   615
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "n-3"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   18
         Top             =   2280
         Width           =   615
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "n-2"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   17
         Top             =   1920
         Width           =   615
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "n-1"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   16
         Top             =   1560
         Width           =   615
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Signo"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   15
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "n"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   14
         Top             =   1200
         Width           =   615
      End
   End
   Begin VB.Frame fraGeneralDatos 
      Caption         =   "Datos"
      ForeColor       =   &H8000000D&
      Height          =   1575
      Left            =   240
      TabIndex        =   6
      Top             =   720
      Width           =   6255
      Begin VB.TextBox TBZona 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
         Height          =   345
         Left            =   3840
         TabIndex        =   9
         Top             =   960
         Width           =   2175
      End
      Begin VB.TextBox TBTickerYahoo 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
         Height          =   345
         Left            =   960
         TabIndex        =   8
         Top             =   960
         Width           =   2055
      End
      Begin VB.TextBox TBNombre 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
         Height          =   345
         Left            =   960
         TabIndex        =   7
         Top             =   480
         Width           =   5055
      End
      Begin VB.Label Label8 
         Caption         =   "Zona"
         Height          =   375
         Left            =   3120
         TabIndex        =   12
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label9 
         Caption         =   "Ticker"
         Height          =   375
         Left            =   240
         TabIndex        =   11
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label14 
         Caption         =   "Nombre"
         Height          =   375
         Left            =   240
         TabIndex        =   10
         Top             =   480
         Width           =   615
      End
   End
   Begin MSComctlLib.TabStrip TBValor 
      Height          =   4080
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7800
      _ExtentX        =   13758
      _ExtentY        =   7197
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   8
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "General"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "MM"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Volumen"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Velas"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "SoportesyResistencias"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab6 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Tendencias"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab7 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Indicadores"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab8 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Fechas"
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
   Begin VB.Frame fraFechasCalendario 
      Caption         =   "Control de fechas con cotización"
      Height          =   3375
      Left            =   240
      TabIndex        =   188
      Top             =   600
      Width           =   7575
      Begin MSComCtl2.MonthView MonthView1 
         Height          =   2820
         Left            =   0
         TabIndex        =   189
         Top             =   360
         Width           =   7530
         _ExtentX        =   13282
         _ExtentY        =   4974
         _Version        =   393216
         ForeColor       =   -2147483630
         BackColor       =   -2147483633
         Appearance      =   1
         MonthColumns    =   3
         ShowToday       =   0   'False
         StartOfWeek     =   50003970
         CurrentDate     =   40428
      End
   End
End
Attribute VB_Name = "frmFichaValor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BExportarWeka_Click()

'ExportarDatosWeka_RSI (VarTickerFichaValor)
'ExportarDatosWeka_Lane (VarTickerFichaValor)

End Sub

Private Sub Form_Load()

Dim ctlControl As Object

For Each ctlControl In frmFichaValor.Controls
    
    If Left(ctlControl.Name, 3) = "fra" Then
    
       If InStr(1, ctlControl.Name, "fraGeneral", vbTextCompare) <> 0 Then
       
          ctlControl.Visible = True
        
       Else
       
          ctlControl.Visible = False
       
       End If
    
    End If
    
Next ctlControl
 

   
TBNombre.Text = ValorNombre
TBTickerYahoo.Text = ValorTickerYahoo
TBZona.Text = ValorZona

' Si la Zona esta en blanco entendemos que es una acción
If ValorZona = "" Then

   fraGeneralControles.Visible = False
   
   Label8.Visible = False
   TBZona.Visible = False

End If

If ValorControlHis Then ChControlHis.Value = 1 Else ChControlHis.Value = 0
If ValorControlCot Then ChControlCot.Value = 1 Else ChControlCot.Value = 0
If ValorControlValorHis Then ChControlValorHis.Value = 1 Else ChControlValorHis.Value = 0
If ValorControlValorCot Then ChControlValorCot.Value = 1 Else ChControlValorCot.Value = 0

If MatrizTemporal(1, 2, 1) = "+" Then

   TBSignoMM20.Text = "Alcista"
   TBSignoMM20.BackColor = &H80FF80
   
Else

   TBSignoMM20.Text = "Bajista"
   TBSignoMM20.BackColor = &H8080FF

End If

If MatrizTemporal(1, 3, 1) = "+" Then

   TBSignoMM50.Text = "Alcista"
   TBSignoMM50.BackColor = &H80FF80
   
Else

   TBSignoMM50.Text = "Bajista"
   TBSignoMM50.BackColor = &H8080FF

End If

If MatrizTemporal(1, 4, 1) = "+" Then

   TBSignoMM200.Text = "Alcista"
   TBSignoMM200.BackColor = &H80FF80
   
Else

   TBSignoMM200.Text = "Bajista"
   TBSignoMM200.BackColor = &H8080FF

End If


For i = 0 To 4

    TBMM20(i).Text = MatrizTemporal(1, 2, i + 2)
    TBMM50(i).Text = MatrizTemporal(1, 3, i + 2)
    TBMM200(i).Text = MatrizTemporal(1, 4, i + 2)
    
    TBK(i).Text = MatrizTemporal(1, 20, i + 1)
    TBD(i).Text = MatrizTemporal(1, 21, i + 1)
    TBDS(i).Text = MatrizTemporal(1, 22, i + 1)
    TBDSS(i).Text = MatrizTemporal(1, 23, i + 1)
    TBRSI(i).Text = MatrizTemporal(1, 24, i + 1)

Next

For i = 0 To 2

    TBVol20(i).Text = MatrizTemporal(1, 5, i + 1)
    TBVol50(i).Text = MatrizTemporal(1, 6, i + 1)
    TBVol200(i).Text = MatrizTemporal(1, 7, i + 1)
    
    TBVela20(i).Text = MatrizTemporal(1, 8, i + 1)
    TBVela50(i).Text = MatrizTemporal(1, 9, i + 1)
    TBVela200(i).Text = MatrizTemporal(1, 10, i + 1)

Next

For i = 0 To 3

    TBSop20(i).Text = MatrizTemporal(1, 11, i + 1)
    TBSop50(i).Text = MatrizTemporal(1, 12, i + 1)
    TBSop200(i).Text = MatrizTemporal(1, 13, i + 1)
    
Next

For i = 0 To 6

    TBTAL(i).Text = MatrizTemporal(1, 14, i + 1)
    TBTAM(i).Text = MatrizTemporal(1, 16, i + 1)
    TBTAC(i).Text = MatrizTemporal(1, 18, i + 1)
    
    TBTBL(i).Text = MatrizTemporal(1, 15, i + 1)
    TBTBM(i).Text = MatrizTemporal(1, 17, i + 1)
    TBTBC(i).Text = MatrizTemporal(1, 19, i + 1)

Next

If InStr(1, MatrizTemporal(1, 14, 8), "ALCISTA", vbTextCompare) <> 0 Then

   LTBTAL.ForeColor = &HFF00&
   
ElseIf InStr(1, MatrizTemporal(1, 14, 8), "BAJISTA", vbTextCompare) <> 0 Then

   LTBTBL.ForeColor = &HFF&
   
ElseIf InStr(1, MatrizTemporal(1, 14, 8), "LATERAL", vbTextCompare) <> 0 Then

   LTBTAL.ForeColor = &HFF0000
   LTBTBL.ForeColor = &HFF0000
   
ElseIf InStr(1, MatrizTemporal(1, 14, 8), "LATERAL-ALCISTA", vbTextCompare) <> 0 Then

   LTBTAL.ForeColor = &H80FF80
   LTBTBL.ForeColor = &H80FF80
   
ElseIf InStr(1, MatrizTemporal(1, 14, 8), "LATERAL-BAJISTA", vbTextCompare) <> 0 Then

   LTBTAL.ForeColor = &H8080FF
   LTBTBL.ForeColor = &H8080FF
   
End If

If InStr(1, MatrizTemporal(1, 16, 8), "ALCISTA", vbTextCompare) <> 0 Then

   LTBTAM.ForeColor = &HFF00&
   
ElseIf InStr(1, MatrizTemporal(1, 16, 8), "BAJISTA", vbTextCompare) <> 0 Then

   LTBTBM.ForeColor = &HFF&
   
ElseIf InStr(1, MatrizTemporal(1, 16, 8), "LATERAL", vbTextCompare) <> 0 Then

   LTBTAM.ForeColor = &HFF0000
   LTBTBM.ForeColor = &HFF0000
   
ElseIf InStr(1, MatrizTemporal(1, 16, 8), "LATERAL-ALCISTA", vbTextCompare) <> 0 Then

   LTBTAM.ForeColor = &H80FF80
   LTBTBM.ForeColor = &H80FF80
   
ElseIf InStr(1, MatrizTemporal(1, 16, 8), "LATERAL-BAJISTA", vbTextCompare) <> 0 Then

   LTBTAM.ForeColor = &H8080FF
   LTBTBM.ForeColor = &H8080FF
   
End If

If InStr(1, MatrizTemporal(1, 18, 8), "ALCISTA", vbTextCompare) <> 0 Then

   LTBTAC.ForeColor = &HFF00&
   
ElseIf InStr(1, MatrizTemporal(1, 18, 8), "BAJISTA", vbTextCompare) <> 0 Then

   LTBTBC.ForeColor = &HFF&
   
ElseIf InStr(1, MatrizTemporal(1, 18, 8), "LATERAL", vbTextCompare) <> 0 Then

   LTBTAC.ForeColor = &HFF0000
   LTBTBC.ForeColor = &HFF0000
   
ElseIf InStr(1, MatrizTemporal(1, 18, 8), "LATERAL-ALCISTA", vbTextCompare) <> 0 Then

   LTBTAC.ForeColor = &H80FF80
   LTBTBC.ForeColor = &H80FF80
   
ElseIf InStr(1, MatrizTemporal(1, 18, 8), "LATERAL-BAJISTA", vbTextCompare) <> 0 Then

   LTBTAC.ForeColor = &H8080FF
   LTBTBC.ForeColor = &H8080FF
   
End If

'&H000000FF& rojo intenso
'&H008080FF& rojo
'&H0000FF00& verde intenso
'&H0080FF80& verde
'&H00FF0000& azul

End Sub

Private Sub MonthView1_DateClick(ByVal DateClicked As Date)

If (MonthView1.DayBold(DateClicked) = True) Then

   MonthView1.DayBold(DateClicked) = False
   
Else

   MonthView1.DayBold(DateClicked) = True

End If

End Sub

Private Sub TBValor_Click()

Dim ctlControl As Object

For Each ctlControl In frmFichaValor.Controls
    
    If Left(ctlControl.Name, 3) = "fra" Then
    
       If InStr(1, ctlControl.Name, "fra" & TBValor.SelectedItem, vbTextCompare) <> 0 Then
       
          ctlControl.Visible = True
        
       Else
       
          ctlControl.Visible = False
       
       End If
    
    End If
    
Next ctlControl

If TBValor.SelectedItem = "Fechas" Then

   MonthView1.Value = Date

End If


' Si la Zona esta en blanco entendemos que es una acción
If ValorZona = "" Then

   fraGeneralControles.Visible = False

End If

End Sub
