VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7275
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9360
   LinkTopic       =   "Form1"
   ScaleHeight     =   7275
   ScaleWidth      =   9360
   StartUpPosition =   3  'Windows Default
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   855
      Left            =   6600
      Top             =   1200
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   1508
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\estudiante\Desktop\Nueva carpeta\autos.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\estudiante\Desktop\Nueva carpeta\autos.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Tabla1"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Command6"
      Height          =   615
      Left            =   5400
      TabIndex        =   13
      Top             =   6000
      Width           =   1695
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Command5"
      Height          =   495
      Left            =   1440
      TabIndex        =   12
      Top             =   6240
      Width           =   3495
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
      Height          =   495
      Left            =   6960
      TabIndex        =   11
      Top             =   5040
      Width           =   1815
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   735
      Left            =   4080
      TabIndex        =   10
      Top             =   5160
      Width           =   2175
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   495
      Left            =   2160
      TabIndex        =   9
      Top             =   5400
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   615
      Left            =   240
      TabIndex        =   8
      Top             =   5280
      Width           =   1455
   End
   Begin VB.TextBox Text4 
      DataField       =   "Direccion"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   2400
      TabIndex        =   7
      Text            =   "Text4"
      Top             =   3360
      Width           =   3735
   End
   Begin VB.TextBox Text3 
      DataField       =   "Precio"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   2520
      TabIndex        =   6
      Text            =   "Text3"
      Top             =   2400
      Width           =   3495
   End
   Begin VB.TextBox Text2 
      DataField       =   "Nombre"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   2520
      TabIndex        =   5
      Text            =   "Text2"
      Top             =   1440
      Width           =   3375
   End
   Begin VB.TextBox Text1 
      DataField       =   "Codigo"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   2520
      TabIndex        =   4
      Top             =   600
      Width           =   3495
   End
   Begin VB.Label Label5 
      Caption         =   "Label5"
      DataField       =   "Fotografia"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   6240
      TabIndex        =   14
      Top             =   4440
      Width           =   2415
   End
   Begin VB.Image Image1 
      DataField       =   "Fotografia"
      DataSource      =   "Adodc1"
      Height          =   1575
      Left            =   6480
      Stretch         =   -1  'True
      Top             =   2640
      Width           =   2055
   End
   Begin VB.Label Label4 
      Caption         =   "Label4"
      Height          =   735
      Left            =   240
      TabIndex        =   3
      Top             =   3240
      Width           =   1815
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      Height          =   615
      Left            =   240
      TabIndex        =   2
      Top             =   2280
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   615
      Left            =   240
      TabIndex        =   1
      Top             =   1440
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   615
      Left            =   360
      TabIndex        =   0
      Top             =   480
      Width           =   1815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Adodc1.Recordset.AddNew
End Sub

Private Sub Command2_Click()
Adodc1.Recordset.MovePrevious
If Adodc1.Recordset.BOF Then
Adodc1.Recordset.MoveNext

End If
End Sub

Private Sub Command3_Click()
Adodc1.Recordset.MoveNext
If Adodc1.Recordset.EOF Then
Adodc1.Recordset.MovePrevious
End If
End Sub

Private Sub Command4_Click()
Adodc1.Recordset.MoveLast
End Sub

Private Sub Command5_Click()
Adodc1.Recordset.Update
End Sub

Private Sub Command6_Click()
Adodc1.Recordset.Delete
End Sub

