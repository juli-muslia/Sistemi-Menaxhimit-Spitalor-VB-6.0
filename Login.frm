VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Login 
   BackColor       =   &H00808000&
   Caption         =   "Hyr"
   ClientHeight    =   7905
   ClientLeft      =   6270
   ClientTop       =   2460
   ClientWidth     =   9090
   LinkTopic       =   "Form1"
   ScaleHeight     =   7905
   ScaleWidth      =   9090
   Begin VB.TextBox mbiemr 
      Height          =   375
      Left            =   840
      TabIndex        =   7
      Top             =   1800
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox emr 
      Height          =   495
      Left            =   840
      TabIndex        =   6
      Top             =   1080
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox id 
      Height          =   375
      Left            =   840
      TabIndex        =   5
      Top             =   600
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   3720
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   3720
      Width           =   2655
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3720
      TabIndex        =   3
      Top             =   2880
      Width           =   2655
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   360
      Top             =   5880
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
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
      Connect         =   $"Login.frx":0000
      OLEDBString     =   $"Login.frx":009F
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from Login"
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
   Begin VB.CommandButton Command1 
      Caption         =   "Hyr"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6360
      TabIndex        =   2
      Top             =   5880
      Width           =   1935
   End
   Begin VB.Label Label2 
      BackColor       =   &H00808000&
      Caption         =   "Fjalekalimi"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1440
      TabIndex        =   1
      Top             =   3720
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackColor       =   &H00808000&
      Caption         =   "Perdoruesi"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1440
      TabIndex        =   0
      Top             =   2880
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   960
      Left            =   3960
      Picture         =   "Login.frx":013E
      Top             =   840
      Width           =   960
   End
End
Attribute VB_Name = "Login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Con As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim strconnect As String
Dim strng As String

Dim User As String
Dim Pass As String
Dim idmj As String
Dim emrmj As String
Dim mbiemrmj As String

Private Sub Command1_Click()
Con = "Provider=MSDASQL.1;Data Source=Juli;Initial Catalog=SQLEXPRESS"
Con.Open "Driver=SQL Server;Server=.\SQLEXPRESS;Database=Hospital;uid=;pwd=;"
Adodc1.RecordSource = "SELECT Login.IdPerdoruesi,Login.IdRoli, Login.IdMjeku, Login.Perdoruesi, Login.Fjalekalimi, Mjeku.Emri, Mjeku.Mbiemri From Login LEFT JOIN Mjeku ON Login.IdMjeku=Mjeku.IdMjek where Perdoruesi = '" & Text1.Text & "' And Fjalekalimi = '" & Text2.Text & "' "
 
 
 
Adodc1.Refresh
User = Text1.Text
Pass = Text2.Text

If Text1.Text = "" Or Text2.Text = "" Then
'Or Text1.Text <> Adodc1.Recordset.Fields("Perdoruesi") Or Text2.Text <> Adodc1.Recordset.Fields("Fjalekalimi") Then
MsgBox "Ju lutem plotesoni fushat !!", vbCritical


Else
If (Adodc1.Recordset.EOF = False) Then
If (Text1.Text = Adodc1.Recordset.Fields("Perdoruesi")) Then
If (Text2.Text = Adodc1.Recordset.Fields("Fjalekalimi")) Then
If Adodc1.Recordset.Fields("IdRoli") = "2" Then
id.Text = Adodc1.Recordset.Fields("IdMjeku")
emr.Text = Adodc1.Recordset.Fields("Emri")
mbiemr.Text = Adodc1.Recordset.Fields("Mbiemri")

Login.Visible = False
Mjeku.Show
Mjeku.SSTab1.Tab = 0
Mjeku.SSTab2.Tab = 0
Mjeku.Text13.Text = Login.id.Text
Mjeku.Label17.Caption = Login.emr.Text
Mjeku.Label18.Caption = Login.mbiemr.Text
Else
If Adodc1.Recordset.Fields("IdRoli") = "1" Then
Unload Me
Admin.Show
Admin.SSTab1.Tab = 0
Admin.SSTab2.Tab = 0
Else
If (Adodc1.Recordset.EOF = False) And Text1.Text <> Adodc1.Recordset.Fields("Perdoruesi") Or Text2.Text <> Adodc1.Recordset.Fields("Fjalekalimi") Then
MsgBox "Perdoruesi ose fjalekalimi jane gabim !! ", vbCritical
End If
End If
End If
End If
End If
End If
End If


Con.Close

End Sub

Private Sub Form_Load()
Text1.Text = ""
Text2.Text = ""
End Sub




