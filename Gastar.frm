VERSION 5.00
Begin VB.Form Gastar 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "GASTOS / INGRESOS"
   ClientHeight    =   3330
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6195
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3330
   ScaleWidth      =   6195
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fringresos 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Concepto de Ingresos"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3255
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   6135
      Begin VB.CommandButton Command3 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Archivar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1560
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   2520
         Width           =   1455
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cancelar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4200
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   2520
         Width           =   1455
      End
      Begin VB.TextBox txtaño2 
         DataSource      =   "Data1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4920
         MaxLength       =   4
         TabIndex        =   21
         Top             =   600
         Width           =   735
      End
      Begin VB.TextBox txtmes2 
         DataSource      =   "Data1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4440
         MaxLength       =   2
         TabIndex        =   20
         Top             =   600
         Width           =   375
      End
      Begin VB.TextBox txtdia2 
         DataSource      =   "Data1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3960
         MaxLength       =   2
         TabIndex        =   19
         Top             =   600
         Width           =   375
      End
      Begin VB.TextBox txtimporte2 
         DataSource      =   "Data1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1560
         TabIndex        =   18
         Top             =   1800
         Width           =   1455
      End
      Begin VB.TextBox txtconcepto2 
         DataSource      =   "Data1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1560
         TabIndex        =   17
         Top             =   1200
         Width           =   4095
      End
      Begin VB.TextBox txtndecaja2 
         DataSource      =   "Data1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1560
         TabIndex        =   16
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "EJ: 200,00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   3120
         TabIndex        =   31
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3240
         TabIndex        =   27
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Importe  :   $ "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   26
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Concepto:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   25
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Nº de Caja:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   24
         Top             =   600
         Width           =   1095
      End
   End
   Begin VB.Frame frgastos 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Concepto de Gastos"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6135
      Begin VB.TextBox txtfecha 
         Height          =   285
         Left            =   4320
         TabIndex        =   14
         Top             =   1800
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Data Data1 
         Caption         =   "Data1"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   495
         Left            =   240
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "Librodecuentas"
         Top             =   2760
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Archivar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2400
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   2520
         Width           =   1455
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cancelar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4320
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   2520
         Width           =   1455
      End
      Begin VB.TextBox txtaño 
         DataSource      =   "Data1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5040
         MaxLength       =   4
         TabIndex        =   10
         Top             =   720
         Width           =   735
      End
      Begin VB.TextBox txtmes 
         DataSource      =   "Data1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4560
         MaxLength       =   2
         TabIndex        =   9
         Top             =   720
         Width           =   375
      End
      Begin VB.TextBox txtdia 
         DataSource      =   "Data1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4080
         MaxLength       =   2
         TabIndex        =   8
         Top             =   720
         Width           =   375
      End
      Begin VB.TextBox txtimporte 
         DataSource      =   "Data1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1680
         TabIndex        =   6
         Top             =   1920
         Width           =   1455
      End
      Begin VB.TextBox txtconcepto 
         DataSource      =   "Data1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1680
         TabIndex        =   5
         Top             =   1320
         Width           =   4095
      End
      Begin VB.TextBox txtndecaja 
         DataSource      =   "Data1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1680
         TabIndex        =   4
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "EJ: 200,00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   3240
         TabIndex        =   13
         Top             =   1920
         Width           =   1095
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3360
         TabIndex        =   7
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Importe  :   $ "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   3
         Top             =   1920
         Width           =   1335
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Concepto:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   2
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Nº de Caja:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   1
         Top             =   720
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Elija una Opcion"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   240
      TabIndex        =   28
      Top             =   480
      Width           =   5895
      Begin VB.CommandButton Command6 
         BackColor       =   &H00C0C0C0&
         Caption         =   "INGRESAR"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   3360
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   840
         Width           =   2295
      End
      Begin VB.CommandButton Command5 
         BackColor       =   &H00C0C0C0&
         Caption         =   "GASTAR"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   840
         Width           =   2295
      End
   End
End
Attribute VB_Name = "Gastar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Data1.DatabaseName = App.Path & ("\VENTAS.mdb")
Data1.RecordSource = "SELECT * FROM Librodecuentas order by Ndecaja"
Data1.Refresh
txtfecha.Text = txtdia.Text & "/" & txtmes.Text & "/" & txtaño.Text
 txtimporte.Text = Format(txtimporte.Text, "$##,##0.00")
 ' If txtcodigo.Text = "" Then
  'MsgBox "Se olvidó de completar el casillero de CODIGO", vbCritical, "CLIENTES - ERROR"
  'txtcodigo.SetFocus
'ElseIf txtcuil.Text = "" Then
 ' MsgBox "Se olvidó de completar el casillero de CUIL", vbCritical, "CLIENTES - ERROR"
  'txtcuil.SetFocus
'ElseIf txtempresa.Text = "" Then
 ' MsgBox "Se olvidó de completar el casillero de EMPRESA", vbCritical, "CLIENTES - ERROR"
  'txtempresa.SetFocus
'ElseIf txtdireccion.Text = "" Then
 ' MsgBox "Se olvidó de completar el casillero de DIRECCION", vbCritical, "CLIENTES - ERROR"
  'txtdireccion.SetFocus
  'Else
Data1.Recordset.AddNew
With Data1
.Recordset.Fields("Ndecaja").Value = txtndecaja.Text
.Recordset.Fields("Fecha").Value = txtfecha.Text

.Recordset.Fields("Concepto").Value = UCase(txtconcepto.Text)
.Recordset.Fields("Gastos").Value = txtimporte.Text
.Refresh
End With

txtndecaja.Text = ""
txtdia.Text = ""
txtmes.Text = ""
txtaño.Text = ""
txtconcepto.Text = ""
txtimporte.Text = ""
txtndecaja.SetFocus
txtdia.Text = Format(Date, "DD")
txtmes.Text = Format(Date, "MM")
txtaño.Text = Format(Date, "YYYY")
End Sub

Private Sub Command2_Click()
txtndecaja.Text = ""
txtdia.Text = ""
txtmes.Text = ""
txtaño.Text = ""
txtconcepto.Text = ""
txtimporte.Text = ""
frgastos.Visible = False
End Sub

Private Sub Command3_Click()
Data1.DatabaseName = App.Path & ("\VENTAS.mdb")
Data1.RecordSource = "SELECT * FROM Librodecuentas order by Ndecaja"
Data1.Refresh
txtfecha.Text = txtdia.Text & "/" & txtmes.Text & "/" & txtaño.Text
 txtimporte2.Text = Format(txtimporte2.Text, "$##,##0.00")
 ' If txtcodigo.Text = "" Then
  'MsgBox "Se olvidó de completar el casillero de CODIGO", vbCritical, "CLIENTES - ERROR"
  'txtcodigo.SetFocus
'ElseIf txtcuil.Text = "" Then
 ' MsgBox "Se olvidó de completar el casillero de CUIL", vbCritical, "CLIENTES - ERROR"
  'txtcuil.SetFocus
'ElseIf txtempresa.Text = "" Then
 ' MsgBox "Se olvidó de completar el casillero de EMPRESA", vbCritical, "CLIENTES - ERROR"
  'txtempresa.SetFocus
'ElseIf txtdireccion.Text = "" Then
 ' MsgBox "Se olvidó de completar el casillero de DIRECCION", vbCritical, "CLIENTES - ERROR"
  'txtdireccion.SetFocus
  'Else
Data1.Recordset.AddNew
With Data1
.Recordset.Fields("Ndecaja").Value = txtndecaja2.Text
.Recordset.Fields("Fecha").Value = txtfecha.Text

.Recordset.Fields("Concepto").Value = UCase(txtconcepto2.Text)
.Recordset.Fields("Ingresos").Value = txtimporte2.Text
.Refresh
End With

txtndecaja2.Text = ""
txtdia2.Text = ""
txtmes2.Text = ""
txtaño2.Text = ""
txtconcepto2.Text = ""
txtimporte2.Text = ""
txtndecaja2.SetFocus
txtdia2.Text = Format(Date, "DD")
txtmes2.Text = Format(Date, "MM")
txtaño2.Text = Format(Date, "YYYY")
End Sub

Private Sub Command4_Click()
txtndecaja2.Text = ""
txtdia2.Text = ""
txtmes2.Text = ""
txtaño2.Text = ""
txtconcepto2.Text = ""
txtimporte2.Text = ""
fringresos.Visible = False
End Sub

Private Sub Command5_Click()
frgastos.Visible = True
fringresos.Visible = False

End Sub

Private Sub Command6_Click()
fringresos.Visible = True
frgastos.Visible = False

End Sub

Private Sub Form_Load()
Data1.DatabaseName = App.Path & ("\VENTAS.mdb")
Data1.RecordSource = "SELECT * FROM Librodecuentas order by Ndecaja"


Data1.Refresh



txtdia.Text = Format(Date, "DD")
txtmes.Text = Format(Date, "MM")
txtaño.Text = Format(Date, "YYYY")
txtdia2.Text = Format(Date, "DD")
txtmes2.Text = Format(Date, "MM")
txtaño2.Text = Format(Date, "YYYY")
frgastos.Visible = False
fringresos.Visible = False
End Sub

