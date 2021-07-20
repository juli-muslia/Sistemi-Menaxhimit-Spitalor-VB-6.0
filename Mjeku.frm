VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Mjeku 
   BackColor       =   &H00808000&
   Caption         =   "Hospital Management System - Mjeku"
   ClientHeight    =   11610
   ClientLeft      =   120
   ClientTop       =   750
   ClientWidth     =   21360
   LinkTopic       =   "Form2"
   ScaleHeight     =   11610
   ScaleMode       =   0  'User
   ScaleWidth      =   21360
   Begin VB.CommandButton Command15 
      Caption         =   "Dil"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   19560
      Picture         =   "Mjeku.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   153
      Top             =   0
      Width           =   1335
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   13800
      TabIndex        =   48
      Top             =   120
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
      _Version        =   393216
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   79626243
      CurrentDate     =   42595
   End
   Begin VB.TextBox Text13 
      Height          =   285
      Left            =   1800
      TabIndex        =   47
      Text            =   "Text13"
      Top             =   600
      Visible         =   0   'False
      Width           =   1455
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   10215
      Left            =   600
      TabIndex        =   0
      Top             =   1200
      Width           =   20175
      _ExtentX        =   35586
      _ExtentY        =   18018
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Pacienti"
      TabPicture(0)   =   "Mjeku.frx":4321
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "SSTab2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Vizitat"
      TabPicture(1)   =   "Mjeku.frx":433D
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "SSTab3"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Fatura"
      TabPicture(2)   =   "Mjeku.frx":4359
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Command27"
      Tab(2).Control(1)=   "Command26"
      Tab(2).Control(2)=   "Command25"
      Tab(2).Control(3)=   "Text40"
      Tab(2).Control(4)=   "Text39"
      Tab(2).Control(5)=   "Text38"
      Tab(2).Control(6)=   "Combo25"
      Tab(2).Control(7)=   "Combo24"
      Tab(2).Control(8)=   "Combo23"
      Tab(2).Control(9)=   "Command24"
      Tab(2).Control(10)=   "Command23"
      Tab(2).Control(11)=   "Command22"
      Tab(2).Control(12)=   "Command21"
      Tab(2).Control(13)=   "DTPicker4"
      Tab(2).Control(14)=   "Text30"
      Tab(2).Control(15)=   "Text29"
      Tab(2).Control(16)=   "Text28"
      Tab(2).Control(17)=   "Text27"
      Tab(2).Control(18)=   "Text26"
      Tab(2).Control(19)=   "Text25"
      Tab(2).Control(20)=   "Text24"
      Tab(2).Control(21)=   "Frame11"
      Tab(2).Control(22)=   "Adodc2"
      Tab(2).Control(23)=   "DataGrid2"
      Tab(2).Control(24)=   "Frame12"
      Tab(2).Control(25)=   "Label80"
      Tab(2).Control(26)=   "Label79"
      Tab(2).Control(27)=   "Label69"
      Tab(2).Control(28)=   "Label45"
      Tab(2).Control(29)=   "Label44"
      Tab(2).Control(30)=   "Label43"
      Tab(2).Control(31)=   "Label42"
      Tab(2).Control(32)=   "Label41"
      Tab(2).Control(33)=   "Label40"
      Tab(2).Control(34)=   "Label39"
      Tab(2).ControlCount=   35
      TabCaption(3)   =   "Raport"
      TabPicture(3)   =   "Mjeku.frx":4375
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Command29"
      Tab(3).Control(1)=   "Command28"
      Tab(3).ControlCount=   2
      Begin VB.CommandButton Command29 
         Caption         =   "Gjenero Raport Fatura"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2415
         Left            =   -65040
         Picture         =   "Mjeku.frx":4391
         Style           =   1  'Graphical
         TabIndex        =   202
         Top             =   3240
         Width           =   2415
      End
      Begin VB.CommandButton Command28 
         Caption         =   "Gjenero Raport Pacienti"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2415
         Left            =   -67680
         Picture         =   "Mjeku.frx":997B
         Style           =   1  'Graphical
         TabIndex        =   201
         Top             =   3240
         Width           =   2415
      End
      Begin VB.CommandButton Command27 
         Caption         =   "Ruaj dhe Printo"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   -66120
         Picture         =   "Mjeku.frx":EF65
         Style           =   1  'Graphical
         TabIndex        =   197
         Top             =   600
         Width           =   1935
      End
      Begin VB.CommandButton Command26 
         Caption         =   "Hiqe"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   -62760
         Picture         =   "Mjeku.frx":12E9C
         Style           =   1  'Graphical
         TabIndex        =   191
         Top             =   8520
         Width           =   1215
      End
      Begin VB.CommandButton Command25 
         Caption         =   "Kerkim i ri "
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   -68280
         Picture         =   "Mjeku.frx":16D2D
         Style           =   1  'Graphical
         TabIndex        =   189
         Top             =   600
         Width           =   1935
      End
      Begin VB.TextBox Text40 
         Height          =   285
         Left            =   -68040
         TabIndex        =   188
         Text            =   "Text40"
         Top             =   9600
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox Text39 
         Height          =   285
         Left            =   -71280
         TabIndex        =   187
         Text            =   "Text39"
         Top             =   9600
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox Text38 
         Height          =   285
         Left            =   -74040
         TabIndex        =   183
         Text            =   "Text38"
         Top             =   9600
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.ComboBox Combo25 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   -67440
         TabIndex        =   182
         Text            =   "Combo25"
         Top             =   6720
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.ComboBox Combo24 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   -70320
         TabIndex        =   181
         Text            =   "Combo24"
         Top             =   6720
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.ComboBox Combo23 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   -73200
         TabIndex        =   180
         Text            =   "Combo23"
         Top             =   6720
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.CommandButton Command24 
         Caption         =   "Shto"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   -64440
         Picture         =   "Mjeku.frx":1B0A8
         Style           =   1  'Graphical
         TabIndex        =   179
         Top             =   8520
         Width           =   1215
      End
      Begin VB.CommandButton Command23 
         Caption         =   "Injeksione"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2295
         Left            =   -67440
         Picture         =   "Mjeku.frx":1F3BC
         Style           =   1  'Graphical
         TabIndex        =   177
         Top             =   4200
         Width           =   2055
      End
      Begin VB.CommandButton Command22 
         Caption         =   "Analiza"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2295
         Left            =   -70320
         Picture         =   "Mjeku.frx":24AE1
         Style           =   1  'Graphical
         TabIndex        =   176
         Top             =   4200
         Width           =   2055
      End
      Begin VB.CommandButton Command21 
         Caption         =   "Ilaçe"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2295
         Left            =   -73200
         Picture         =   "Mjeku.frx":2AB19
         Style           =   1  'Graphical
         TabIndex        =   172
         Top             =   4200
         Width           =   2055
      End
      Begin MSComCtl2.DTPicker DTPicker4 
         Height          =   375
         Left            =   -64560
         TabIndex        =   152
         Top             =   1080
         Visible         =   0   'False
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         _Version        =   393216
         Format          =   79626241
         CurrentDate     =   42599
      End
      Begin TabDlg.SSTab SSTab3 
         Height          =   8895
         Left            =   -70560
         TabIndex        =   137
         Top             =   840
         Width           =   11295
         _ExtentX        =   19923
         _ExtentY        =   15690
         _Version        =   393216
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "Shto Vizite"
         TabPicture(0)   =   "Mjeku.frx":304A8
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Label61"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Label62"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "Label63"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "Label64"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "Label66"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "Text32"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "Combo19"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).Control(7)=   "Combo20"
         Tab(0).Control(7).Enabled=   0   'False
         Tab(0).Control(8)=   "Command11"
         Tab(0).Control(8).Enabled=   0   'False
         Tab(0).Control(9)=   "Command12"
         Tab(0).Control(9).Enabled=   0   'False
         Tab(0).Control(10)=   "Command13"
         Tab(0).Control(10).Enabled=   0   'False
         Tab(0).Control(11)=   "Command14"
         Tab(0).Control(11).Enabled=   0   'False
         Tab(0).Control(12)=   "Adodc3"
         Tab(0).Control(12).Enabled=   0   'False
         Tab(0).Control(13)=   "DataGrid3"
         Tab(0).Control(13).Enabled=   0   'False
         Tab(0).Control(14)=   "Text33"
         Tab(0).Control(14).Enabled=   0   'False
         Tab(0).Control(15)=   "DTPicker5"
         Tab(0).Control(15).Enabled=   0   'False
         Tab(0).ControlCount=   16
         TabCaption(1)   =   "Vizitat e mia"
         TabPicture(1)   =   "Mjeku.frx":304C4
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Label68"
         Tab(1).Control(1)=   "Label70"
         Tab(1).Control(2)=   "Label71"
         Tab(1).Control(3)=   "Label72"
         Tab(1).Control(4)=   "Label73"
         Tab(1).Control(5)=   "Label74"
         Tab(1).Control(6)=   "Label65"
         Tab(1).Control(7)=   "Command16"
         Tab(1).Control(8)=   "Command17"
         Tab(1).Control(9)=   "Command18"
         Tab(1).Control(10)=   "Command19"
         Tab(1).Control(11)=   "Command20"
         Tab(1).Control(12)=   "Text35"
         Tab(1).Control(13)=   "Combo21"
         Tab(1).Control(14)=   "Combo22"
         Tab(1).Control(15)=   "Text36"
         Tab(1).Control(16)=   "Text37"
         Tab(1).Control(17)=   "DataGrid4"
         Tab(1).Control(18)=   "Adodc4"
         Tab(1).Control(19)=   "DTPicker6"
         Tab(1).Control(20)=   "Combo26"
         Tab(1).ControlCount=   21
         Begin VB.ComboBox Combo26 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            ItemData        =   "Mjeku.frx":304E0
            Left            =   -70200
            List            =   "Mjeku.frx":304EA
            TabIndex        =   196
            Top             =   5160
            Width           =   2415
         End
         Begin MSComCtl2.DTPicker DTPicker6 
            Height          =   375
            Left            =   -70200
            TabIndex        =   190
            Top             =   4680
            Width           =   2895
            _ExtentX        =   5106
            _ExtentY        =   661
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "dd/MM/yyyy HH:mm:ss"
            Format          =   79626243
            CurrentDate     =   42608
         End
         Begin MSAdodcLib.Adodc Adodc4 
            Height          =   495
            Left            =   -74880
            Top             =   5400
            Visible         =   0   'False
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   873
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
            Connect         =   $"Mjeku.frx":30502
            OLEDBString     =   $"Mjeku.frx":305A1
            OLEDBFile       =   ""
            DataSourceName  =   ""
            OtherAttributes =   ""
            UserName        =   ""
            Password        =   ""
            RecordSource    =   $"Mjeku.frx":30640
            Caption         =   "Adodc4"
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
         Begin MSDataGridLib.DataGrid DataGrid4 
            Bindings        =   "Mjeku.frx":306D7
            Height          =   2055
            Left            =   -74880
            TabIndex        =   171
            Top             =   6120
            Width           =   11050
            _ExtentX        =   19500
            _ExtentY        =   3625
            _Version        =   393216
            AllowUpdate     =   0   'False
            HeadLines       =   1
            RowHeight       =   15
            FormatLocked    =   -1  'True
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnCount     =   7
            BeginProperty Column00 
               DataField       =   "IdVizita"
               Caption         =   "IdVizita"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column01 
               DataField       =   "IdPacienti"
               Caption         =   "IdPacienti"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column02 
               DataField       =   "Emri"
               Caption         =   "Emri"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column03 
               DataField       =   "Atesia"
               Caption         =   "Atesia"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column04 
               DataField       =   "Mbiemri"
               Caption         =   "Mbiemri"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column05 
               DataField       =   "DateVizita"
               Caption         =   "DateVizita"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column06 
               DataField       =   "Status"
               Caption         =   "Status"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            SplitCount      =   1
            BeginProperty Split0 
               BeginProperty Column00 
                  ColumnWidth     =   915.024
               EndProperty
               BeginProperty Column01 
                  ColumnWidth     =   1739.906
               EndProperty
               BeginProperty Column02 
                  ColumnWidth     =   1739.906
               EndProperty
               BeginProperty Column03 
                  ColumnWidth     =   1739.906
               EndProperty
               BeginProperty Column04 
                  ColumnWidth     =   1739.906
               EndProperty
               BeginProperty Column05 
                  ColumnWidth     =   1739.906
               EndProperty
               BeginProperty Column06 
                  ColumnWidth     =   1739.906
               EndProperty
            EndProperty
         End
         Begin VB.TextBox Text37 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   -72960
            TabIndex        =   170
            Top             =   1560
            Width           =   975
         End
         Begin VB.TextBox Text36 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   -70200
            TabIndex        =   168
            Top             =   4200
            Width           =   2415
         End
         Begin VB.ComboBox Combo22 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   -70200
            TabIndex        =   167
            Top             =   3720
            Width           =   2415
         End
         Begin VB.ComboBox Combo21 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   -70200
            TabIndex        =   166
            Top             =   3240
            Width           =   2415
         End
         Begin VB.TextBox Text35 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   -70200
            TabIndex        =   165
            Top             =   2760
            Width           =   2415
         End
         Begin VB.CommandButton Command20 
            Caption         =   "Mbyll"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   -68040
            Picture         =   "Mjeku.frx":306EC
            Style           =   1  'Graphical
            TabIndex        =   159
            Top             =   600
            Width           =   855
         End
         Begin VB.CommandButton Command19 
            Caption         =   "Refresh"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   -69000
            Picture         =   "Mjeku.frx":33DCE
            Style           =   1  'Graphical
            TabIndex        =   158
            Top             =   600
            Width           =   855
         End
         Begin VB.CommandButton Command18 
            Caption         =   "Fshij"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   -69960
            Picture         =   "Mjeku.frx":3720F
            Style           =   1  'Graphical
            TabIndex        =   157
            Top             =   600
            Width           =   855
         End
         Begin VB.CommandButton Command17 
            Caption         =   "RiRuaj"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   -70920
            Picture         =   "Mjeku.frx":3A4E2
            Style           =   1  'Graphical
            TabIndex        =   156
            Top             =   600
            Width           =   855
         End
         Begin VB.CommandButton Command16 
            Caption         =   "I Ri"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   -71880
            Picture         =   "Mjeku.frx":3D6A9
            Style           =   1  'Graphical
            TabIndex        =   155
            Top             =   600
            Width           =   855
         End
         Begin MSComCtl2.DTPicker DTPicker5 
            Height          =   375
            Left            =   4440
            TabIndex        =   154
            Top             =   3960
            Width           =   2895
            _ExtentX        =   5106
            _ExtentY        =   661
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "dd/MM/yyyy HH:mm:ss"
            Format          =   79626243
            CurrentDate     =   42600.5
         End
         Begin VB.TextBox Text33 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   4440
            TabIndex        =   151
            Top             =   3450
            Width           =   2535
         End
         Begin MSDataGridLib.DataGrid DataGrid3 
            Bindings        =   "Mjeku.frx":40994
            Height          =   2055
            Left            =   1320
            TabIndex        =   149
            Top             =   5400
            Width           =   8415
            _ExtentX        =   14843
            _ExtentY        =   3625
            _Version        =   393216
            AllowUpdate     =   0   'False
            HeadLines       =   1
            RowHeight       =   15
            FormatLocked    =   -1  'True
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnCount     =   4
            BeginProperty Column00 
               DataField       =   "IdPacienti"
               Caption         =   "IdPacienti"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column01 
               DataField       =   "Emri"
               Caption         =   "Emri"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column02 
               DataField       =   "Atesia"
               Caption         =   "Atesia"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column03 
               DataField       =   "Mbiemri"
               Caption         =   "Mbiemri"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            SplitCount      =   1
            BeginProperty Split0 
               BeginProperty Column00 
                  ColumnWidth     =   1739.906
               EndProperty
               BeginProperty Column01 
                  ColumnWidth     =   1739.906
               EndProperty
               BeginProperty Column02 
                  ColumnWidth     =   1739.906
               EndProperty
               BeginProperty Column03 
                  ColumnWidth     =   1739.906
               EndProperty
            EndProperty
         End
         Begin MSAdodcLib.Adodc Adodc3 
            Height          =   375
            Left            =   240
            Top             =   4680
            Visible         =   0   'False
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   661
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
            Connect         =   $"Mjeku.frx":409A9
            OLEDBString     =   $"Mjeku.frx":40A48
            OLEDBFile       =   ""
            DataSourceName  =   ""
            OtherAttributes =   ""
            UserName        =   ""
            Password        =   ""
            RecordSource    =   "Select IdPacienti,Emri,Atesia,Mbiemri From Pacienti"
            Caption         =   "Adodc3"
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
         Begin VB.CommandButton Command14 
            Caption         =   "Mbyll"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   6840
            Picture         =   "Mjeku.frx":40AE7
            Style           =   1  'Graphical
            TabIndex        =   148
            Top             =   600
            Width           =   855
         End
         Begin VB.CommandButton Command13 
            Caption         =   "RiRuaj"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   5880
            Picture         =   "Mjeku.frx":441C9
            Style           =   1  'Graphical
            TabIndex        =   147
            Top             =   600
            Width           =   855
         End
         Begin VB.CommandButton Command12 
            Caption         =   "Ruaj"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   4920
            Picture         =   "Mjeku.frx":47390
            Style           =   1  'Graphical
            TabIndex        =   146
            Top             =   600
            Width           =   855
         End
         Begin VB.CommandButton Command11 
            Caption         =   "I Ri"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   3960
            Picture         =   "Mjeku.frx":4A408
            Style           =   1  'Graphical
            TabIndex        =   145
            Top             =   600
            Width           =   855
         End
         Begin VB.ComboBox Combo20 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   4440
            TabIndex        =   144
            Top             =   3000
            Width           =   2535
         End
         Begin VB.ComboBox Combo19 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   4440
            TabIndex        =   143
            Top             =   2520
            Width           =   2535
         End
         Begin VB.TextBox Text32 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   4440
            TabIndex        =   142
            Top             =   2040
            Width           =   2535
         End
         Begin VB.Label Label65 
            Caption         =   "Statusi i Vizites"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   -72400
            TabIndex        =   195
            Top             =   5160
            Width           =   2025
         End
         Begin VB.Label Label74 
            Caption         =   "Kodi i Vizites"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   -74880
            TabIndex        =   169
            Top             =   1560
            Width           =   1815
         End
         Begin VB.Label Label73 
            Caption         =   "ID e Pacientit"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   -72120
            TabIndex        =   164
            Top             =   2760
            Width           =   1815
         End
         Begin VB.Label Label72 
            Caption         =   "Emri i Pacientit"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   -72375
            TabIndex        =   163
            Top             =   3240
            Width           =   2055
         End
         Begin VB.Label Label71 
            Caption         =   "Mbiemri i Pacientit"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   -72840
            TabIndex        =   162
            Top             =   3720
            Width           =   2415
         End
         Begin VB.Label Label70 
            Caption         =   "Data dhe Ora e re e Vizites"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   -73770
            TabIndex        =   161
            Top             =   4680
            Width           =   3360
         End
         Begin VB.Label Label68 
            Caption         =   "Atesia e Pacientit"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   -72555
            TabIndex        =   160
            Top             =   4200
            Width           =   2295
         End
         Begin VB.Label Label66 
            Caption         =   "Atesia e Pacientit"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2085
            TabIndex        =   150
            Top             =   3480
            Width           =   2295
         End
         Begin VB.Label Label64 
            Caption         =   "Data dhe Ora Vizites"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1635
            TabIndex        =   141
            Top             =   3960
            Width           =   2655
         End
         Begin VB.Label Label63 
            Caption         =   "Mbiemri i Pacientit"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1800
            TabIndex        =   140
            Top             =   3000
            Width           =   2415
         End
         Begin VB.Label Label62 
            Caption         =   "Emri i Pacientit"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2265
            TabIndex        =   139
            Top             =   2520
            Width           =   2055
         End
         Begin VB.Label Label61 
            Caption         =   "ID e Pacientit"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2520
            TabIndex        =   138
            Top             =   2040
            Width           =   1815
         End
      End
      Begin VB.TextBox Text30 
         Height          =   285
         Left            =   -65880
         TabIndex        =   119
         Top             =   1320
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox Text29 
         Height          =   285
         Left            =   -65880
         TabIndex        =   112
         Top             =   1080
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox Text28 
         Height          =   285
         Left            =   -65880
         TabIndex        =   111
         Top             =   840
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox Text27 
         Height          =   285
         Left            =   -67800
         TabIndex        =   110
         Top             =   1560
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox Text26 
         Height          =   285
         Left            =   -67800
         TabIndex        =   109
         Top             =   1320
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox Text25 
         Height          =   285
         Left            =   -67800
         TabIndex        =   108
         Top             =   1080
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox Text24 
         Height          =   285
         Left            =   -67800
         TabIndex        =   107
         Top             =   840
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Frame Frame11 
         Caption         =   "Kerko Pacient"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   -74760
         TabIndex        =   101
         Top             =   720
         Width           =   5895
         Begin VB.ComboBox Combo17 
            Height          =   315
            Left            =   2640
            TabIndex        =   103
            Top             =   240
            Width           =   2655
         End
         Begin VB.ComboBox Combo18 
            Enabled         =   0   'False
            Height          =   315
            Left            =   2640
            TabIndex        =   102
            Top             =   600
            Width           =   2655
         End
         Begin VB.Label Label36 
            Caption         =   "Emri i Pacientit"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   600
            TabIndex        =   105
            Top             =   240
            Width           =   2055
         End
         Begin VB.Label Label35 
            Caption         =   "Mbiemri i Pacientit"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   104
            Top             =   600
            Width           =   2535
         End
      End
      Begin MSAdodcLib.Adodc Adodc2 
         Height          =   330
         Left            =   -74160
         Top             =   360
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
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
         Connect         =   $"Mjeku.frx":4D6F3
         OLEDBString     =   $"Mjeku.frx":4D792
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   $"Mjeku.frx":4D831
         Caption         =   "Adodc2"
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
      Begin MSDataGridLib.DataGrid DataGrid2 
         Bindings        =   "Mjeku.frx":4D8E4
         Height          =   2055
         Left            =   -74760
         TabIndex        =   100
         Top             =   2040
         Width           =   11775
         _ExtentX        =   20770
         _ExtentY        =   3625
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   15
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   8
         BeginProperty Column00 
            DataField       =   "IdPacienti"
            Caption         =   "IdPacienti"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "Emri"
            Caption         =   "Emri"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column02 
            DataField       =   "Atesia"
            Caption         =   "Atesia"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column03 
            DataField       =   "Mbiemri"
            Caption         =   "Mbiemri"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column04 
            DataField       =   "Gjinia"
            Caption         =   "Gjinia"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column05 
            DataField       =   "Datelindja"
            Caption         =   "Datelindja"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column06 
            DataField       =   "Vendlindja"
            Caption         =   "Vendlindja"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column07 
            DataField       =   "Kontakt"
            Caption         =   "Kontakt"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
               ColumnWidth     =   1739.906
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1739.906
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1739.906
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1739.906
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   464.882
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   1140.095
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   1739.906
            EndProperty
            BeginProperty Column07 
               ColumnWidth     =   1140.095
            EndProperty
         EndProperty
      End
      Begin VB.Frame Frame12 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fature"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   9375
         Left            =   -62880
         TabIndex        =   99
         Top             =   480
         Width           =   7815
         Begin VB.TextBox Text14 
            Height          =   3975
            Left            =   3840
            MultiLine       =   -1  'True
            TabIndex        =   200
            Text            =   "Mjeku.frx":4D8F9
            Top             =   3960
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.TextBox Text31 
            Height          =   3975
            Left            =   2880
            MultiLine       =   -1  'True
            TabIndex        =   199
            Text            =   "Mjeku.frx":4D900
            Top             =   3960
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.ListBox List2 
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3960
            Left            =   4800
            TabIndex        =   194
            Top             =   3960
            Width           =   2535
         End
         Begin VB.ListBox List1 
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3960
            Left            =   240
            TabIndex        =   193
            Top             =   3960
            Width           =   2535
         End
         Begin VB.CommandButton Command10 
            Caption         =   "Afisho Faturen"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   1920
            TabIndex        =   178
            Top             =   8520
            Width           =   1575
         End
         Begin VB.Frame Frame13 
            Caption         =   "Te Dhenat e Pacientit"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2055
            Left            =   240
            TabIndex        =   121
            Top             =   1080
            Width           =   7335
            Begin VB.Label Label59 
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   1200
               TabIndex        =   135
               Top             =   1440
               Width           =   2535
            End
            Begin VB.Label Label58 
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   1200
               TabIndex        =   134
               Top             =   1080
               Width           =   2535
            End
            Begin VB.Label Label57 
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   5520
               TabIndex        =   133
               Top             =   1080
               Width           =   1935
            End
            Begin VB.Label Label56 
               Caption         =   "Mosha : "
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   4740
               TabIndex        =   132
               Top             =   1080
               Width           =   855
            End
            Begin VB.Label Label55 
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   5520
               TabIndex        =   131
               Top             =   720
               Width           =   1815
            End
            Begin VB.Label Label54 
               Caption         =   "Datelindja :"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   4440
               TabIndex        =   130
               Top             =   720
               Width           =   1095
            End
            Begin VB.Label Label53 
               Caption         =   "Mbiemri : "
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   375
               TabIndex        =   129
               Top             =   1440
               Width           =   855
            End
            Begin VB.Label Label52 
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   1200
               TabIndex        =   128
               Top             =   720
               Width           =   2535
            End
            Begin VB.Label Label51 
               Caption         =   "Atesia :"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   525
               TabIndex        =   127
               Top             =   1080
               Width           =   735
            End
            Begin VB.Label Label49 
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   5520
               TabIndex        =   126
               Top             =   360
               Width           =   1815
            End
            Begin VB.Label Label48 
               Caption         =   "Gjinia : "
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   4800
               TabIndex        =   125
               Top             =   360
               Width           =   735
            End
            Begin VB.Label Label47 
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   1200
               TabIndex        =   124
               Top             =   360
               Width           =   2535
            End
            Begin VB.Label Label38 
               Caption         =   "Emri :"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   645
               TabIndex        =   123
               Top             =   720
               Width           =   615
            End
            Begin VB.Label Label50 
               Caption         =   "Id Pacienti :"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Left            =   120
               TabIndex        =   122
               Top             =   360
               Width           =   1095
            End
         End
         Begin VB.Label Label67 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Lek"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   6600
            TabIndex        =   198
            Top             =   8640
            Width           =   495
         End
         Begin VB.Label Label46 
            BackColor       =   &H00C0C0C0&
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   5640
            TabIndex        =   192
            Top             =   8640
            Width           =   975
         End
         Begin VB.Label Label77 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Çmimi"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   5160
            TabIndex        =   175
            Top             =   3360
            Width           =   1455
         End
         Begin VB.Label Label76 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Produkti"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   240
            TabIndex        =   174
            Top             =   3360
            Width           =   2535
         End
         Begin VB.Label Label75 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Çmimi Total:"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3960
            TabIndex        =   173
            Top             =   8640
            Width           =   1455
         End
         Begin VB.Label Label37 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   240
            TabIndex        =   106
            Top             =   360
            Width           =   7335
         End
      End
      Begin TabDlg.SSTab SSTab2 
         Height          =   9135
         Left            =   720
         TabIndex        =   1
         Top             =   720
         Width           =   18735
         _ExtentX        =   33046
         _ExtentY        =   16113
         _Version        =   393216
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "Shto Pacient"
         TabPicture(0)   =   "Mjeku.frx":4D907
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Command1"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Command2"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "Command3"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "Command4"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "Frame1"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "Frame2"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "Frame3"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).Control(7)=   "Frame4"
         Tab(0).Control(7).Enabled=   0   'False
         Tab(0).Control(8)=   "Frame5"
         Tab(0).Control(8).Enabled=   0   'False
         Tab(0).Control(9)=   "Command30"
         Tab(0).Control(9).Enabled=   0   'False
         Tab(0).ControlCount=   10
         TabCaption(1)   =   "Ndrysho Pacient"
         TabPicture(1)   =   "Mjeku.frx":4D923
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Command5"
         Tab(1).Control(1)=   "Command6"
         Tab(1).Control(2)=   "Command7"
         Tab(1).Control(3)=   "Command8"
         Tab(1).Control(4)=   "Command9"
         Tab(1).Control(5)=   "DataGrid1"
         Tab(1).Control(6)=   "Frame6"
         Tab(1).Control(7)=   "Adodc1"
         Tab(1).Control(8)=   "Frame7"
         Tab(1).Control(9)=   "Frame8"
         Tab(1).Control(10)=   "Frame9"
         Tab(1).Control(11)=   "Frame10"
         Tab(1).ControlCount=   12
         Begin VB.CommandButton Command30 
            Caption         =   "Gjenero Id Pacienti"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1920
            TabIndex        =   203
            Top             =   1320
            Width           =   2535
         End
         Begin VB.Frame Frame10 
            Caption         =   "Injeksionet"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1095
            Left            =   -65880
            TabIndex        =   86
            Top             =   6600
            Width           =   9375
            Begin VB.TextBox Text23 
               Height          =   285
               Left            =   6240
               TabIndex        =   98
               Text            =   "Text23"
               Top             =   120
               Visible         =   0   'False
               Width           =   735
            End
            Begin VB.TextBox Text22 
               Height          =   285
               Left            =   1200
               TabIndex        =   97
               Text            =   "Text22"
               Top             =   120
               Visible         =   0   'False
               Width           =   1335
            End
            Begin VB.ComboBox Combo16 
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   435
               Left            =   7200
               TabIndex        =   94
               Top             =   480
               Width           =   1695
            End
            Begin VB.ComboBox Combo15 
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   435
               Left            =   2640
               TabIndex        =   93
               Top             =   480
               Width           =   2055
            End
            Begin VB.Label Label34 
               Caption         =   "Tipi i Injeksionit"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   15.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   4920
               TabIndex        =   90
               Top             =   480
               Width           =   2175
            End
            Begin VB.Label Label33 
               Caption         =   "Emri i Injeksionit"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   15.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   240
               TabIndex        =   89
               Top             =   480
               Width           =   2295
            End
         End
         Begin VB.Frame Frame9 
            Caption         =   "Analiza"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1095
            Left            =   -65880
            TabIndex        =   85
            Top             =   5520
            Width           =   9375
            Begin VB.TextBox Text21 
               Height          =   285
               Left            =   7560
               TabIndex        =   96
               Text            =   "Text21"
               Top             =   360
               Visible         =   0   'False
               Width           =   975
            End
            Begin VB.ComboBox Combo14 
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   435
               Left            =   3000
               TabIndex        =   92
               Top             =   480
               Width           =   3495
            End
            Begin VB.Label Label32 
               Caption         =   "Emri i Analizes"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   15.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   240
               TabIndex        =   88
               Top             =   480
               Width           =   2055
            End
         End
         Begin VB.Frame Frame8 
            Caption         =   "Patologjia"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1095
            Left            =   -65880
            TabIndex        =   84
            Top             =   4440
            Width           =   9375
            Begin VB.TextBox Text20 
               Height          =   285
               Left            =   7560
               TabIndex        =   95
               Text            =   "Text20"
               Top             =   360
               Visible         =   0   'False
               Width           =   855
            End
            Begin VB.ComboBox Combo13 
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   435
               Left            =   3000
               TabIndex        =   91
               Top             =   360
               Width           =   3375
            End
            Begin VB.Label Label31 
               Caption         =   "Emri i Patologjise"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   15.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   240
               TabIndex        =   87
               Top             =   405
               Width           =   2295
            End
         End
         Begin VB.Frame Frame7 
            Caption         =   "Gjeneralitetet"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3255
            Left            =   -74880
            TabIndex        =   63
            Top             =   4440
            Width           =   8655
            Begin VB.ComboBox Combo12 
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   435
               ItemData        =   "Mjeku.frx":4D93F
               Left            =   6120
               List            =   "Mjeku.frx":4D949
               TabIndex        =   83
               Top             =   2520
               Width           =   2175
            End
            Begin VB.ComboBox Combo11 
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   435
               ItemData        =   "Mjeku.frx":4D95B
               Left            =   6120
               List            =   "Mjeku.frx":4D965
               TabIndex        =   82
               Top             =   2040
               Width           =   2175
            End
            Begin VB.TextBox Text19 
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   405
               Left            =   6120
               TabIndex        =   81
               Top             =   1560
               Width           =   2175
            End
            Begin VB.ComboBox Combo10 
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   435
               ItemData        =   "Mjeku.frx":4D97F
               Left            =   6120
               List            =   "Mjeku.frx":4D9EF
               TabIndex        =   80
               Top             =   1080
               Width           =   2175
            End
            Begin VB.ComboBox Combo9 
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   435
               ItemData        =   "Mjeku.frx":4DB23
               Left            =   6120
               List            =   "Mjeku.frx":4DB2D
               TabIndex        =   79
               Top             =   600
               Width           =   720
            End
            Begin MSComCtl2.DTPicker DTPicker3 
               Height          =   375
               Left            =   1680
               TabIndex        =   78
               Top             =   2520
               Width           =   2175
               _ExtentX        =   3836
               _ExtentY        =   661
               _Version        =   393216
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Times New Roman"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               CustomFormat    =   "dd/MM/yyyy"
               Format          =   79626243
               CurrentDate     =   42596
            End
            Begin VB.TextBox Text18 
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   405
               Left            =   1680
               TabIndex        =   77
               Top             =   2040
               Width           =   2175
            End
            Begin VB.TextBox Text17 
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
               Left            =   1680
               TabIndex        =   76
               Top             =   1560
               Width           =   2175
            End
            Begin VB.TextBox Text16 
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
               Left            =   1680
               TabIndex        =   75
               Top             =   1080
               Width           =   2175
            End
            Begin VB.TextBox Text15 
               Enabled         =   0   'False
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
               Left            =   1680
               TabIndex        =   74
               Top             =   600
               Width           =   2175
            End
            Begin VB.Label Label30 
               Caption         =   "Statusi"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   15.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   5040
               TabIndex        =   73
               Top             =   2520
               Width           =   975
            End
            Begin VB.Label Label29 
               Caption         =   "Statusi Civil"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   15.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   4320
               TabIndex        =   72
               Top             =   2040
               Width           =   1575
            End
            Begin VB.Label Label28 
               Caption         =   "Kontakt"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   15.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   4875
               TabIndex        =   71
               Top             =   1560
               Width           =   975
            End
            Begin VB.Label Label27 
               Caption         =   "Vendlindja"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   15.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   4560
               TabIndex        =   70
               Top             =   1080
               Width           =   1455
            End
            Begin VB.Label Label26 
               Caption         =   "Datelindja"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   15.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   240
               TabIndex        =   69
               Top             =   2520
               Width           =   1335
            End
            Begin VB.Label Label25 
               Caption         =   "Gjinia"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   15.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   5100
               TabIndex        =   68
               Top             =   600
               Width           =   855
            End
            Begin VB.Label Label24 
               Caption         =   "Mbiemri"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   15.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   405
               TabIndex        =   67
               Top             =   2040
               Width           =   1095
            End
            Begin VB.Label Label23 
               Caption         =   "Atesia"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   15.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   675
               TabIndex        =   66
               Top             =   1560
               Width           =   855
            End
            Begin VB.Label Label22 
               Caption         =   "Emri"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   15.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   795
               TabIndex        =   65
               Top             =   1080
               Width           =   615
            End
            Begin VB.Label Label21 
               Caption         =   "Id Pacienti"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   15.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   120
               TabIndex        =   64
               Top             =   600
               Width           =   1335
            End
         End
         Begin MSAdodcLib.Adodc Adodc1 
            Height          =   495
            Left            =   -62400
            Top             =   840
            Visible         =   0   'False
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   873
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
            Connect         =   $"Mjeku.frx":4DB37
            OLEDBString     =   $"Mjeku.frx":4DBD6
            OLEDBFile       =   ""
            DataSourceName  =   ""
            OtherAttributes =   ""
            UserName        =   ""
            Password        =   ""
            RecordSource    =   $"Mjeku.frx":4DC75
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
         Begin VB.Frame Frame6 
            Caption         =   "Kerko Pacient"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1095
            Left            =   -74760
            TabIndex        =   58
            Top             =   480
            Width           =   5895
            Begin VB.ComboBox Combo8 
               Enabled         =   0   'False
               Height          =   315
               Left            =   2640
               TabIndex        =   62
               Top             =   600
               Width           =   2655
            End
            Begin VB.ComboBox Combo7 
               Height          =   315
               Left            =   2640
               TabIndex        =   61
               Text            =   "Combo7"
               Top             =   240
               Width           =   2655
            End
            Begin VB.Label Label20 
               Caption         =   "Mbiemri i Pacientit"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   15.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   120
               TabIndex        =   60
               Top             =   600
               Width           =   2535
            End
            Begin VB.Label Label19 
               Caption         =   "Emri i Pacientit"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   15.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   600
               TabIndex        =   59
               Top             =   240
               Width           =   2055
            End
         End
         Begin MSDataGridLib.DataGrid DataGrid1 
            Bindings        =   "Mjeku.frx":4DF9E
            Height          =   2535
            Left            =   -74880
            TabIndex        =   57
            Top             =   1800
            Width           =   18495
            _ExtentX        =   32623
            _ExtentY        =   4471
            _Version        =   393216
            AllowUpdate     =   0   'False
            HeadLines       =   1
            RowHeight       =   15
            FormatLocked    =   -1  'True
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnCount     =   14
            BeginProperty Column00 
               DataField       =   "IdPacienti"
               Caption         =   "IdPacienti"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column01 
               DataField       =   "Emri"
               Caption         =   "Emri"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column02 
               DataField       =   "Atesia"
               Caption         =   "Atesia"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column03 
               DataField       =   "Mbiemri"
               Caption         =   "Mbiemri"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column04 
               DataField       =   "Gjinia"
               Caption         =   "Gjinia"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column05 
               DataField       =   "Datelindja"
               Caption         =   "Datelindja"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column06 
               DataField       =   "Vendlindja"
               Caption         =   "Vendlindja"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column07 
               DataField       =   "Kontakt"
               Caption         =   "Kontakt"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column08 
               DataField       =   "Statusi_Civil"
               Caption         =   "Statusi_Civil"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column09 
               DataField       =   "Statusi"
               Caption         =   "Statusi"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column10 
               DataField       =   "TipiInjeksionit"
               Caption         =   "TipiInjeksionit"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column11 
               DataField       =   "EmerInjeksion"
               Caption         =   "EmerInjeksion"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column12 
               DataField       =   "EmerPatologjia"
               Caption         =   "EmerPatologjia"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column13 
               DataField       =   "EmerAnaliza"
               Caption         =   "EmerAnaliza"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            SplitCount      =   1
            BeginProperty Split0 
               BeginProperty Column00 
                  ColumnWidth     =   1739.906
               EndProperty
               BeginProperty Column01 
                  ColumnWidth     =   1739.906
               EndProperty
               BeginProperty Column02 
                  ColumnWidth     =   1739.906
               EndProperty
               BeginProperty Column03 
                  ColumnWidth     =   1739.906
               EndProperty
               BeginProperty Column04 
                  ColumnWidth     =   464.882
               EndProperty
               BeginProperty Column05 
                  ColumnWidth     =   1140.095
               EndProperty
               BeginProperty Column06 
                  ColumnWidth     =   1739.906
               EndProperty
               BeginProperty Column07 
                  ColumnWidth     =   1140.095
               EndProperty
               BeginProperty Column08 
                  ColumnWidth     =   1739.906
               EndProperty
               BeginProperty Column09 
                  ColumnWidth     =   1739.906
               EndProperty
               BeginProperty Column10 
                  ColumnWidth     =   1739.906
               EndProperty
               BeginProperty Column11 
                  ColumnWidth     =   1739.906
               EndProperty
               BeginProperty Column12 
                  ColumnWidth     =   1739.906
               EndProperty
               BeginProperty Column13 
                  ColumnWidth     =   1739.906
               EndProperty
            EndProperty
         End
         Begin VB.CommandButton Command9 
            Caption         =   "Mbyll"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   -64080
            Picture         =   "Mjeku.frx":4DFB3
            Style           =   1  'Graphical
            TabIndex        =   56
            Top             =   600
            Width           =   855
         End
         Begin VB.CommandButton Command8 
            Caption         =   "Refresh"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   -65040
            Picture         =   "Mjeku.frx":51695
            Style           =   1  'Graphical
            TabIndex        =   55
            Top             =   600
            Width           =   855
         End
         Begin VB.CommandButton Command7 
            Caption         =   "Fshij"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   -66000
            Picture         =   "Mjeku.frx":54AD6
            Style           =   1  'Graphical
            TabIndex        =   54
            Top             =   600
            Width           =   855
         End
         Begin VB.CommandButton Command6 
            Caption         =   "RiRuaj"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   -66960
            Picture         =   "Mjeku.frx":57DA9
            Style           =   1  'Graphical
            TabIndex        =   53
            Top             =   600
            Width           =   855
         End
         Begin VB.CommandButton Command5 
            Caption         =   "I Ri"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   -67920
            Picture         =   "Mjeku.frx":5AF70
            Style           =   1  'Graphical
            TabIndex        =   52
            Top             =   600
            Width           =   855
         End
         Begin VB.Frame Frame5 
            Caption         =   "Receta"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2415
            Left            =   240
            TabIndex        =   35
            Top             =   6480
            Width           =   10455
            Begin VB.TextBox Text12 
               Height          =   1935
               Left            =   120
               MultiLine       =   -1  'True
               TabIndex        =   45
               Top             =   360
               Width           =   10215
            End
         End
         Begin VB.Frame Frame4 
            Caption         =   "Analiza"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1575
            Left            =   10920
            TabIndex        =   34
            Top             =   4800
            Width           =   7455
            Begin VB.TextBox Text11 
               Height          =   285
               Left            =   6240
               TabIndex        =   44
               Text            =   "Text11"
               Top             =   600
               Visible         =   0   'False
               Width           =   735
            End
            Begin VB.ComboBox Combo6 
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   435
               Left            =   2760
               TabIndex        =   43
               Top             =   600
               Width           =   3135
            End
            Begin VB.Label Label14 
               Caption         =   "Emri i Analizes"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   15.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   480
               TabIndex        =   42
               Top             =   600
               Width           =   1935
            End
         End
         Begin VB.Frame Frame3 
            Caption         =   "Injeksionet"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1575
            Left            =   240
            TabIndex        =   33
            Top             =   4800
            Width           =   10455
            Begin VB.TextBox Text10 
               Height          =   285
               Left            =   6000
               TabIndex        =   41
               Text            =   "Text10"
               Top             =   840
               Visible         =   0   'False
               Width           =   855
            End
            Begin VB.TextBox Text9 
               Height          =   285
               Left            =   6120
               TabIndex        =   40
               Text            =   "Text9"
               Top             =   360
               Visible         =   0   'False
               Width           =   735
            End
            Begin VB.ComboBox Combo5 
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   435
               Left            =   2760
               TabIndex        =   39
               Top             =   840
               Width           =   2895
            End
            Begin VB.ComboBox Combo4 
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   435
               Left            =   2760
               TabIndex        =   38
               Top             =   360
               Width           =   2895
            End
            Begin VB.Label Label13 
               Caption         =   "Tipi i Injeksionit"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   15.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   360
               TabIndex        =   37
               Top             =   840
               Width           =   2175
            End
            Begin VB.Label Label12 
               Caption         =   "Emri i Injeksionit"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   15.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   240
               TabIndex        =   36
               Top             =   360
               Width           =   2295
            End
         End
         Begin VB.Frame Frame2 
            Caption         =   "Patologjia"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3135
            Left            =   10920
            TabIndex        =   27
            Top             =   1560
            Width           =   7455
            Begin VB.TextBox Text8 
               Height          =   405
               Left            =   6120
               TabIndex        =   32
               Text            =   "Text8"
               Top             =   840
               Visible         =   0   'False
               Width           =   975
            End
            Begin VB.ComboBox Combo3 
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   435
               Left            =   3000
               TabIndex        =   31
               Top             =   480
               Width           =   3015
            End
            Begin VB.TextBox Text7 
               Height          =   1575
               Left            =   120
               Locked          =   -1  'True
               MultiLine       =   -1  'True
               OLEDragMode     =   1  'Automatic
               TabIndex        =   29
               Top             =   1440
               Width           =   7215
            End
            Begin VB.Label Label11 
               Caption         =   "Pershkrimi i Patologjise"
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
               Left            =   120
               TabIndex        =   30
               Top             =   1080
               Width           =   3015
            End
            Begin VB.Label Label10 
               Caption         =   "Emer Patologjie"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   15.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   720
               TabIndex        =   28
               Top             =   480
               Width           =   2175
            End
         End
         Begin VB.Frame Frame1 
            Caption         =   "Gjeneralietet"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3135
            Left            =   240
            TabIndex        =   6
            Top             =   1560
            Width           =   10455
            Begin VB.ComboBox Combo2 
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   435
               ItemData        =   "Mjeku.frx":5E25B
               Left            =   6000
               List            =   "Mjeku.frx":5E265
               TabIndex        =   26
               Text            =   "Selekto"
               Top             =   2400
               Width           =   2535
            End
            Begin VB.ComboBox Combo1 
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   435
               ItemData        =   "Mjeku.frx":5E27F
               Left            =   6000
               List            =   "Mjeku.frx":5E2EF
               TabIndex        =   25
               Text            =   "Selekto"
               Top             =   1440
               Width           =   2535
            End
            Begin VB.TextBox Text6 
               Height          =   285
               Left            =   8160
               TabIndex        =   24
               Top             =   960
               Visible         =   0   'False
               Width           =   1095
            End
            Begin VB.OptionButton Option2 
               Caption         =   "F"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   15.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   7080
               TabIndex        =   23
               Top             =   960
               Width           =   735
            End
            Begin VB.OptionButton Option1 
               Caption         =   "M"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   15.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   6120
               TabIndex        =   22
               Top             =   960
               Width           =   735
            End
            Begin VB.TextBox Text5 
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
               Left            =   6000
               TabIndex        =   21
               Top             =   1920
               Width           =   2535
            End
            Begin MSComCtl2.DTPicker DTPicker1 
               Height          =   375
               Left            =   6000
               TabIndex        =   20
               Top             =   480
               Width           =   2535
               _ExtentX        =   4471
               _ExtentY        =   661
               _Version        =   393216
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Times New Roman"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               CustomFormat    =   "dd/MM/yyyy"
               Format          =   79626243
               CurrentDate     =   42594
            End
            Begin VB.TextBox Text4 
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
               Left            =   1680
               TabIndex        =   19
               Top             =   1920
               Width           =   2535
            End
            Begin VB.TextBox Text3 
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
               Left            =   1680
               TabIndex        =   18
               Top             =   1440
               Width           =   2535
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
               Left            =   1680
               TabIndex        =   17
               Top             =   960
               Width           =   2535
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
               Left            =   1680
               TabIndex        =   16
               Top             =   480
               Width           =   2535
            End
            Begin VB.Label Label1 
               Caption         =   "Id Pacienti"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   15.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   120
               TabIndex        =   15
               Top             =   480
               Width           =   1335
            End
            Begin VB.Label Label2 
               Caption         =   "Emri"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   15.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   840
               TabIndex        =   14
               Top             =   960
               Width           =   615
            End
            Begin VB.Label Label3 
               Caption         =   "Atesia"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   15.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   705
               TabIndex        =   13
               Top             =   1440
               Width           =   855
            End
            Begin VB.Label Label4 
               Caption         =   "Mbiemri"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   15.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   360
               TabIndex        =   12
               Top             =   1920
               Width           =   1095
            End
            Begin VB.Label Label5 
               Caption         =   "Gjinia"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   15.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   5100
               TabIndex        =   11
               Top             =   960
               Width           =   855
            End
            Begin VB.Label Label6 
               Caption         =   "Datelindja"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   15.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Left            =   4560
               TabIndex        =   10
               Top             =   480
               Width           =   1335
            End
            Begin VB.Label Label7 
               Caption         =   "Vendlindja"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   15.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   4500
               TabIndex        =   9
               Top             =   1440
               Width           =   1455
            End
            Begin VB.Label Label8 
               Caption         =   "Kontakt"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   15.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   4875
               TabIndex        =   8
               Top             =   1920
               Width           =   975
            End
            Begin VB.Label Label9 
               Caption         =   "Statusi Civil"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   15.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   4320
               TabIndex        =   7
               Top             =   2400
               Width           =   1575
            End
         End
         Begin VB.CommandButton Command4 
            Caption         =   "Mbyll"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   10440
            Picture         =   "Mjeku.frx":5E423
            Style           =   1  'Graphical
            TabIndex        =   5
            Top             =   600
            Width           =   855
         End
         Begin VB.CommandButton Command3 
            Caption         =   "RiRuaj"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   9480
            Picture         =   "Mjeku.frx":61B05
            Style           =   1  'Graphical
            TabIndex        =   4
            Top             =   600
            Width           =   855
         End
         Begin VB.CommandButton Command2 
            Caption         =   "Ruaj"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   8520
            Picture         =   "Mjeku.frx":64CCC
            Style           =   1  'Graphical
            TabIndex        =   3
            Top             =   600
            Width           =   855
         End
         Begin VB.CommandButton Command1 
            Caption         =   "I Ri"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   7560
            Picture         =   "Mjeku.frx":67D44
            Style           =   1  'Graphical
            TabIndex        =   2
            Top             =   600
            Width           =   855
         End
      End
      Begin VB.Label Label80 
         Caption         =   "cmimi"
         Height          =   255
         Left            =   -68760
         TabIndex        =   186
         Top             =   9600
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label79 
         Caption         =   "cmimi"
         Height          =   255
         Left            =   -71880
         TabIndex        =   185
         Top             =   9600
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label69 
         Caption         =   "cmimi"
         Height          =   255
         Left            =   -74640
         TabIndex        =   184
         Top             =   9600
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label45 
         Caption         =   "Data"
         Height          =   375
         Left            =   -66360
         TabIndex        =   120
         Top             =   1320
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label44 
         Caption         =   "Datelindja"
         Height          =   375
         Left            =   -66600
         TabIndex        =   118
         Top             =   1080
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label Label43 
         Caption         =   "Gjinia"
         Height          =   255
         Left            =   -66360
         TabIndex        =   117
         Top             =   840
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label Label42 
         Caption         =   "Mbiemri"
         Height          =   255
         Left            =   -68640
         TabIndex        =   116
         Top             =   1560
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label Label41 
         Caption         =   "Atesia"
         Height          =   255
         Left            =   -68640
         TabIndex        =   115
         Top             =   1320
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label Label40 
         Caption         =   "Emri"
         Height          =   375
         Left            =   -68400
         TabIndex        =   114
         Top             =   1080
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label39 
         Caption         =   "Id pacienti"
         Height          =   255
         Left            =   -68640
         TabIndex        =   113
         Top             =   840
         Visible         =   0   'False
         Width           =   855
      End
   End
   Begin VB.Label Label60 
      Caption         =   "Label60"
      Height          =   495
      Left            =   9720
      TabIndex        =   136
      Top             =   5520
      Width           =   1215
   End
   Begin VB.Label Label18 
      BackColor       =   &H00808000&
      Caption         =   "Label18"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   7080
      TabIndex        =   51
      Top             =   165
      Width           =   1215
   End
   Begin VB.Label Label17 
      BackColor       =   &H00808000&
      Caption         =   "Label17"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   6000
      TabIndex        =   50
      Top             =   165
      Width           =   1215
   End
   Begin VB.Label Label16 
      BackColor       =   &H00808000&
      Caption         =   "Data:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   12960
      TabIndex        =   49
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label15 
      BackColor       =   &H00808000&
      Caption         =   "Miresevini: Dr."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   3480
      TabIndex        =   46
      Top             =   120
      Width           =   2535
   End
   Begin VB.Menu pacienti 
      Caption         =   "&Pacienti"
      Begin VB.Menu shtoPacient 
         Caption         =   "Shto Pacient"
      End
      Begin VB.Menu ndryshoPacient 
         Caption         =   "Ndrysho Pacient"
      End
      Begin VB.Menu raportPacient 
         Caption         =   "Raport"
      End
   End
   Begin VB.Menu vizitat 
      Caption         =   "&Vizitat"
      Begin VB.Menu shtoVizita 
         Caption         =   "Shto Vizita"
      End
      Begin VB.Menu vizitateMia 
         Caption         =   "Vizitat e Mia"
      End
   End
   Begin VB.Menu fatura 
      Caption         =   "&Fatura"
   End
   Begin VB.Menu raport 
      Caption         =   "&Raport"
   End
End
Attribute VB_Name = "Mjeku"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Deklarimi i variablave per tu lidhur me bazen e te dhenave

Dim Con As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim rec As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim strconnect As String
Dim strng As String
Const RndString = "0123456789abcdefghijklmnopqrstuvwxyz"

Private Function RndWord(Optional Chrlength As Integer = 8) As String
Dim TempWord As String
Dim LoopVar As Integer

For LoopVar = 1 To Chrlength
    TempWord = TempWord & Mid(RndString, Int(Rnd * 16) + 1, 1)
    Next LoopVar
    RndWord = TempWord
End Function

Private Function Rndchr() As String
Rndchr = Mid(RndString, Int(Rnd * 16) + 1, 1)
End Function

Private Sub Command30_Click()
Text1.Text = RndWord(7)
End Sub

'======================================================SHTO PACIENT=======================================================================
Private Sub Option1_Click()
Text6.Text = "M"
End Sub

Private Sub Option2_Click()
Text6.Text = "F"
End Sub
'MBUSHJA E COMBOBOXEVE

Private Sub Combo3_Add()
Con = "Provider=MSDASQL.1;Data Source=Juli;Initial Catalog=SQLEXPRESS"
Con.Open "Driver=SQL Server;Server=.\SQLEXPRESS;Database=Hospital;uid=;pwd=;"
rec.Open "select EmerPatologjia from Patologjia", Con, adOpenUnspecified, adLockReadOnly
Combo3.AddItem "Selekto"
Combo3.Text = Me.Combo3.List(0)
While Not rec.EOF
Combo3.AddItem rec(0)
rec.MoveNext
Wend
Con.Close
End Sub

Private Sub Combo3_Click()
Con = "Provider=MSDASQL.1;Data Source=Juli;Initial Catalog=SQLEXPRESS"
Con.Open "Driver=SQL Server;Server=.\SQLEXPRESS;Database=Hospital;uid=;pwd=;"
rec.Open "select IdPatologjia,Pershkrimi from Patologjia Where EmerPatologjia = '" & Combo3 & "'", Con, adOpenUnspecified, adLockReadOnly
While Not rec.EOF
Text8.Text = rec(0)
Text7.Text = rec(1)
rec.MoveNext
Wend
Con.Close
End Sub

Private Sub Combo4_Add()
Con = "Provider=MSDASQL.1;Data Source=Juli;Initial Catalog=SQLEXPRESS"
Con.Open "Driver=SQL Server;Server=.\SQLEXPRESS;Database=Hospital;uid=;pwd=;"
rec.Open "select EmerInjeksion from Injeksionet", Con, adOpenUnspecified, adLockReadOnly
Combo4.AddItem "Selekto"
Combo4.Text = Me.Combo4.List(0)
While Not rec.EOF
Combo4.AddItem rec(0)
rec.MoveNext
Wend
Con.Close
End Sub

Private Sub Combo4_Click()
Con = "Provider=MSDASQL.1;Data Source=Juli;Initial Catalog=SQLEXPRESS"
Con.Open "Driver=SQL Server;Server=.\SQLEXPRESS;Database=Hospital;uid=;pwd=;"
rec.Open "select IdInjeksion from Injeksionet Where EmerInjeksion = '" & Combo4 & "'", Con, adOpenUnspecified, adLockReadOnly
While Not rec.EOF
Text9.Text = rec(0)
rec.MoveNext
Wend
Con.Close
Combo5.Enabled = True
End Sub

Private Sub Combo5_Add()
Con = "Provider=MSDASQL.1;Data Source=Juli;Initial Catalog=SQLEXPRESS"
Con.Open "Driver=SQL Server;Server=.\SQLEXPRESS;Database=Hospital;uid=;pwd=;"
rec.Open "select TipiInjeksionit from TipInjeksioni", Con, adOpenUnspecified, adLockReadOnly
Combo5.AddItem "Selekto"
Combo5.Text = Me.Combo5.List(0)
While Not rec.EOF
Combo5.AddItem rec(0)
rec.MoveNext
Wend
Con.Close
End Sub

Private Sub Combo5_Click()
Con = "Provider=MSDASQL.1;Data Source=Juli;Initial Catalog=SQLEXPRESS"
Con.Open "Driver=SQL Server;Server=.\SQLEXPRESS;Database=Hospital;uid=;pwd=;"
rec.Open "select IdTipInjeksioni from TipInjeksioni Where TipiInjeksionit = '" & Combo5 & "'", Con, adOpenUnspecified, adLockReadOnly
While Not rec.EOF
Text10.Text = rec(0)
rec.MoveNext
Wend
Con.Close
End Sub

Private Sub Combo6_Add()
Con = "Provider=MSDASQL.1;Data Source=Juli;Initial Catalog=SQLEXPRESS"
Con.Open "Driver=SQL Server;Server=.\SQLEXPRESS;Database=Hospital;uid=;pwd=;"
rec.Open "select EmerAnaliza from Analiza", Con, adOpenUnspecified, adLockReadOnly
Combo6.AddItem "Selekto"
Combo6.Text = Me.Combo6.List(0)
While Not rec.EOF
Combo6.AddItem rec(0)
rec.MoveNext
Wend
Con.Close
End Sub

Private Sub Combo6_Click()
Con = "Provider=MSDASQL.1;Data Source=Juli;Initial Catalog=SQLEXPRESS"
Con.Open "Driver=SQL Server;Server=.\SQLEXPRESS;Database=Hospital;uid=;pwd=;"
rec.Open "select IdAnaliza from Analiza Where EmerAnaliza = '" & Combo6 & "'", Con, adOpenUnspecified, adLockReadOnly
While Not rec.EOF
Text11.Text = rec(0)
rec.MoveNext
Wend
Con.Close
End Sub

'Klikimi i butonit I RI
Private Sub Command1_Click()
Text1.Text = " "
Text2.Text = " "
Text3.Text = " "
Text4.Text = " "
DTPicker1.Value = Format(Now(), "dd/MM/yyyy")
Combo1 = "Selekto"
Text5.Text = " "
Text6.Text = " "
Combo2 = "Selekto"
Combo3 = "Selekto"
Text7.Text = " "
Text8.Text = " "
Combo4 = "Selekto"
Text9.Text = " "
Combo5 = "Selekto"
Text10.Text = " "
Combo6 = "Selekto"
Text11.Text = " "
Text12.Text = " "
End Sub

' Klikimi i butonit RUAJ
Private Sub Command2_Click()
        
Dim status As String
status = "Aktiv"
Con = "Provider=MSDASQL.1;Data Source=Juli;Initial Catalog=SQLEXPRESS"
Con.Open "Driver=SQL Server;Server=.\SQLEXPRESS;Database=Hospital;uid=;pwd=;"
If Text1.Text = "" Or Text2.Text = "" Or Text3.Text = "" Or Text4.Text = "" Or Combo1 = "" Or Text5.Text = "" Or Text6.Text = "" Or Combo2 = "" Or Combo3 = "" Or Text7.Text = "" Or Text8.Text = "" Or Combo4 = "" Or Text9.Text = "" Or Combo5 = "" Or Text10.Text = "" Or Combo6 = "" Or Text11.Text = "" Or Text12.Text = "" Then
MsgBox "PLOTESONI TE GJITHA FUSHAT"

Else
    On Error GoTo ProcError
Set rs = Con.Execute( _
        "INSERT INTO Pacienti(IdPacienti, Emri, Atesia, Mbiemri, Gjinia, Vendlindja, Datelindja, Kontakt, Statusi_Civil,Statusi, DtRegjistrimi)VALUES('" & Text1.Text & "','" & Text2.Text & "','" & Text3.Text & "','" & Text4.Text & "','" & Text6.Text & "','" & Combo1 & "','" & Format$(DTPicker1.Value, "yyyy.mm.dd") & "','" & Text5.Text & "','" & Combo2 & "', '" & status & "','" & Format$(DTPicker2.Value, "yyyy.mm.dd") & "')")

Set rs = Con.Execute( _
        "INSERT INTO Pacient_Patologji(IdPacient,IdPatologji) VALUES('" & Text1.Text & "','" & Text8.Text & "')")
        
Set rs = Con.Execute( _
       "INSERT INTO Pacient_Injeksion(IdPacient,IdMjeku,IdInjeksion,Data,IdTipInjeksioni) VALUES('" & Text1.Text & "','" & Text13.Text & "','" & Text9.Text & "','" & Format$(DTPicker2.Value, "yyyy.mm.dd") & "','" & Text10.Text & "')")
       
Set rs = Con.Execute( _
        "INSERT INTO Analize_Pacient(IdPacient,IdAnalize,IdMjeku,Data) VALUES('" & Text1.Text & "','" & Text11.Text & "','" & Text13.Text & "','" & Format$(DTPicker2.Value, "yyyy.mm.dd") & "')")
  
Set rs = Con.Execute( _
        "INSERT INTO Receta(IdMjeku,IdPacient,Pershkrimi,Data) VALUES('" & Text13.Text & "','" & Text1.Text & "','" & Text12.Text & "','" & Format$(DTPicker2.Value, "yyyy.mm.dd") & "')")
        
MsgBox "RUAJTJA E TE DHENAVE U KRYE ME SUKSES."
End If
Con.Close
ProcExit:
Exit Sub
ProcError:

MsgBox "TE DHENAT NUK U RUAJTEN. SIGUROHUNI QE KENI PLOTESUAR TE GJITHA FUSHAT."
Con.Close
Resume ProcExit
Con.Close
End Sub

' Klikimi i butonit RiRuaj
Private Sub Command3_Click()
Mjeku.SSTab2.Tab = 1
End Sub

'Klikimi i butonit MBYLL

Private Sub Command4_Click()
Unload Me
End Sub

'===============================================FUND SHTO PACIENT ====================================================================================
'===============================================NDRYSHO PACIENT ====================================================================================


'Klikimi i butonit I RI
Private Sub Command5_Click()
Mjeku.SSTab2.Tab = 0
End Sub
'Klikimi i butonit RIRUAJ
Private Sub Command6_Click()
Con = "Provider=MSDASQL.1;Data Source=Juli;Initial Catalog=SQLEXPRESS"
Con.Open "Driver=SQL Server;Server=.\SQLEXPRESS;Database=Hospital;uid=;pwd=;"

If Text15.Text = "" Or Text16.Text = "" Or Text17.Text = "" Or Text18.Text = "" Or DTPicker3.Value = "" Or Combo9.Text = "" Or Combo10.Text = "" Or Text19.Text = "" Or Combo11.Text = "" Or Combo12.Text = "" Or Combo13.Text = "" Or Text20.Text = "" Or Combo14.Text = "" Or Text21.Text = "" Or Combo15.Text = "" Or Text22.Text = "" Or Combo16.Text = "" Or Text23.Text = "" Then
MsgBox "PLOTESO FUSHAT !"
Else
    On Error GoTo ProcError
    
Set rs = Con.Execute("Update Pacienti Set Pacienti.Emri = '" & Text16.Text & "', Pacienti.Atesia = '" & Text17.Text & "', Pacienti.Mbiemri = '" & Text18.Text & "', Pacienti.Gjinia = '" & Combo9.Text & "', Pacienti.Vendlindja = '" & Combo10.Text & "', Pacienti.Datelindja = '" & Format$(DTPicker3.Value, "yyyy.mm.dd") & "', Pacienti.Kontakt = '" & Text19.Text & "', Pacienti.Statusi_Civil = '" & Combo11.Text & "', Pacienti.Statusi = '" & Combo12.Text & "'  Where IdPacienti = '" & Text15.Text & "' ")
 
 Set rs = Con.Execute("Update Analize_Pacient Set Analize_Pacient.IdAnalize = '" & Text21.Text & "' Where IdPacient = '" & Text15.Text & "'")
 
  Set rs = Con.Execute("Update Pacient_Patologji Set Pacient_Patologji.IdPatologji = '" & Text20.Text & "' Where IdPacient = '" & Text15.Text & "'")
  
  Set rs = Con.Execute("Update Pacient_Injeksion Set Pacient_Injeksion.IdInjeksion = '" & Text22.Text & "' , Pacient_Injeksion.IdTipInjeksioni = '" & Text23.Text & "' Where IdPacient = '" & Text15.Text & "'")
 
 
MsgBox "MODIFIKIMI I TE DHENAVE U KRYE ME SUKSES."

End If
Con.Close
ProcExit:
Exit Sub
ProcError:


MsgBox "Error Number: " & Err.Number & " With The Description ->> " & Err.Description & " <<- Occured."

MsgBox "TE DHENAT NUK U MODIFIKUAN."
Con.Close
Resume ProcExit
Con.Close

End Sub
'Klikimi i butonit FSHIJ
Private Sub Command7_Click()
status = "Pasiv"
Con = "Provider=MSDASQL.1;Data Source=Juli;Initial Catalog=SQLEXPRESS"
Con.Open "Driver=SQL Server;Server=.\SQLEXPRESS;Database=Hospital;uid=;pwd=;"
If Text15.Text = "" Then
MsgBox "PLOTESO FUSHAT."
Else
    On Error GoTo ProcError
    
Set rs = Con.Execute("Update Pacienti Set Statusi = '" & status & "' Where IdPacienti = '" & Text15.Text & "' ")
    
 
MsgBox "MODIFIKIMI I TE DHENAVE U KRYE ME SUKSES."

End If
Con.Close
ProcExit:
Exit Sub
ProcError:


MsgBox "Error Number: " & Err.Number & " With The Description ->> " & Err.Description & " <<- Occured."

MsgBox "TE DHENAT NUK U MODIFIKUAN."
Con.Close
Resume ProcExit
Con.Close

End Sub

'Klikimi i butonit REFRESH
Private Sub Command8_Click()
Text15.Text = ""
Text16.Text = ""
Text17.Text = ""
Text18.Text = ""
DTPicker3.Value = Format(Now(), "dd/MM/yyyy")
Combo9.Text = ""
Combo10.Text = ""
Text19.Text = ""
Combo11.Text = ""
Combo12.Text = ""

Combo13.Text = ""
Text20.Text = ""

Combo14.Text = ""
Text21.Text = ""

Combo15.Text = ""
Text22.Text = ""

Combo16.Text = ""
Text23.Text = ""

Combo7.Text = "Selekto"
Combo8.Text = "Selekto"

Adodc1.RecordSource = "Select Distinct Pacienti.IdPacienti, Pacienti.Emri, Pacienti.Atesia, Pacienti.Mbiemri, Pacienti.Gjinia, Pacienti.Datelindja, Pacienti.Vendlindja, Pacienti.Kontakt, Pacienti.Statusi_Civil, Pacienti.Statusi, TipInjeksioni.TipiInjeksionit, Injeksionet.EmerInjeksion, Patologjia.EmerPatologjia, Analiza.EmerAnaliza From Pacienti, Analize_Pacient, Pacient_Patologji, Pacient_Injeksion, Injeksionet, Patologjia,TipInjeksioni, Analiza Where TipInjeksioni.IdTipInjeksioni = Injeksionet.IdTipi And Pacienti.IdPacienti = Analize_Pacient.IdPacient And Pacienti.IdPacienti = Pacient_Patologji.IdPacient And Pacienti.IdPacienti = Pacient_Injeksion.IdPacient And Patologjia.IdPatologjia = Pacient_Patologji.IdPatologji And Analiza.IdAnaliza = Analize_Pacient.IdAnalize And Injeksionet.IdInjeksion = Pacient_Injeksion.IdInjeksion "
Adodc1.Refresh
DataGrid1.Refresh

End Sub

'Klikimi i butonit MBYLL
Private Sub Command9_Click()
Unload Me
End Sub

'Kerkimi me ane te Emrit dhe Mbiemrit te Pacientit

Private Sub Combo7_Add()
Combo7.Clear
Dim strng As String
Con = "Provider=MSDASQL.1;Data Source=Juli;Initial Catalog=SQLEXPRESS"
Con.Open "Driver=SQL Server;Server=.\SQLEXPRESS;Database=Hospital;uid=;pwd=;"
rec.Open "select Emri from Pacienti", Con, adOpenUnspecified, adLockReadOnly
Combo7.AddItem "Selekto"
Combo7.Text = Me.Combo7.List(0)

While Not rec.EOF
Combo7.AddItem rec(0)
rec.MoveNext
Wend
Con.Close
End Sub

Private Sub Combo7_Click()
If Combo7.ListIndex = 0 Then
        Adodc1.RecordSource = "Select Distinct Pacienti.IdPacienti,Pacienti.Emri, Pacienti.Atesia, Pacienti.Mbiemri, Pacienti.Gjinia, Pacienti.Datelindja, Pacienti.Vendlindja, Pacienti.Kontakt, Pacienti.Statusi_Civil, Pacienti.Statusi, TipInjeksioni.TipiInjeksionit, Injeksionet.EmerInjeksion, Patologjia.EmerPatologjia, Analiza.EmerAnaliza From Pacienti, Analize_Pacient, Pacient_Patologji, Pacient_Injeksion, Injeksionet, Patologjia,TipInjeksioni, Analiza Where TipInjeksioni.IdTipInjeksioni = Injeksionet.IdTipi And Pacienti.IdPacienti = Analize_Pacient.IdPacient And Pacienti.IdPacienti = Pacient_Patologji.IdPacient And Pacienti.IdPacienti = Pacient_Injeksion.IdPacient And Patologjia.IdPatologjia = Pacient_Patologji.IdPatologji And Analiza.IdAnaliza = Analize_Pacient.IdAnalize And Injeksionet.IdInjeksion = Pacient_Injeksion.IdInjeksion"
    Else
        Adodc1.RecordSource = "Select Distinct Pacienti.IdPacienti,Pacienti.Emri, Pacienti.Atesia, Pacienti.Mbiemri, Pacienti.Gjinia, Pacienti.Datelindja, Pacienti.Vendlindja, Pacienti.Kontakt, Pacienti.Statusi_Civil, Pacienti.Statusi, TipInjeksioni.TipiInjeksionit, Injeksionet.EmerInjeksion, Patologjia.EmerPatologjia, Analiza.EmerAnaliza From Pacienti, Analize_Pacient, Pacient_Patologji, Pacient_Injeksion, Injeksionet, Patologjia,TipInjeksioni, Analiza Where TipInjeksioni.IdTipInjeksioni = Injeksionet.IdTipi And Pacienti.IdPacienti = Analize_Pacient.IdPacient And Pacienti.IdPacienti = Pacient_Patologji.IdPacient And Pacienti.IdPacienti = Pacient_Injeksion.IdPacient And Patologjia.IdPatologjia = Pacient_Patologji.IdPatologji And Analiza.IdAnaliza = Analize_Pacient.IdAnalize And Injeksionet.IdInjeksion = Pacient_Injeksion.IdInjeksion And Emri = '" & Combo7 & "'"
    End If
        Combo7.Refresh
    Adodc1.Refresh
DataGrid1.Refresh
Combo8.Enabled = True

Con = "Provider=MSDASQL.1;Data Source=Juli;Initial Catalog=SQLEXPRESS"
Con.Open "Driver=SQL Server;Server=.\SQLEXPRESS;Database=Hospital;uid=;pwd=;"
rec.Open "Select Mbiemri from Pacienti Where Emri = '" & Combo7 & "'", Con, adOpenUnspecified, adLockReadOnly

Combo8.AddItem "Selekto"
Combo8.Text = Me.Combo8.List(0)

While Not rec.EOF
Combo8.AddItem rec(0)
rec.MoveNext
Wend
Con.Close

End Sub

Private Sub Combo8_Click()
If Combo8.ListIndex = 0 Then
        Adodc1.RecordSource = "Select Distinct Pacienti.IdPacienti,Pacienti.Emri, Pacienti.Atesia, Pacienti.Mbiemri, Pacienti.Gjinia, Pacienti.Datelindja, Pacienti.Vendlindja, Pacienti.Kontakt, Pacienti.Statusi_Civil, Pacienti.Statusi, TipInjeksioni.TipiInjeksionit, Injeksionet.EmerInjeksion, Patologjia.EmerPatologjia, Analiza.EmerAnaliza From Pacienti, Analize_Pacient, Pacient_Patologji, Pacient_Injeksion, Injeksionet, Patologjia,TipInjeksioni, Analiza Where TipInjeksioni.IdTipInjeksioni = Injeksionet.IdTipi And Pacienti.IdPacienti = Analize_Pacient.IdPacient And Pacienti.IdPacienti = Pacient_Patologji.IdPacient And Pacienti.IdPacienti = Pacient_Injeksion.IdPacient And Patologjia.IdPatologjia = Pacient_Patologji.IdPatologji And Analiza.IdAnaliza = Analize_Pacient.IdAnalize And Injeksionet.IdInjeksion = Pacient_Injeksion.IdInjeksion"
    Else
               Adodc1.RecordSource = "Select Distinct Pacienti.IdPacienti,Pacienti.Emri, Pacienti.Atesia, Pacienti.Mbiemri, Pacienti.Gjinia, Pacienti.Datelindja, Pacienti.Vendlindja, Pacienti.Kontakt, Pacienti.Statusi_Civil, Pacienti.Statusi, TipInjeksioni.TipiInjeksionit, Injeksionet.EmerInjeksion, Patologjia.EmerPatologjia, Analiza.EmerAnaliza From Pacienti, Analize_Pacient, Pacient_Patologji, Pacient_Injeksion, Injeksionet, Patologjia,TipInjeksioni, Analiza Where TipInjeksioni.IdTipInjeksioni = Injeksionet.IdTipi And Pacienti.IdPacienti = Analize_Pacient.IdPacient And Pacienti.IdPacienti = Pacient_Patologji.IdPacient And Pacienti.IdPacienti = Pacient_Injeksion.IdPacient And Patologjia.IdPatologjia = Pacient_Patologji.IdPatologji And Analiza.IdAnaliza = Analize_Pacient.IdAnalize And Injeksionet.IdInjeksion = Pacient_Injeksion.IdInjeksion And Emri = '" & Combo7 & "' And Mbiemri = '" & Combo8 & "'"
    End If
    
    Combo8.Refresh
    Adodc1.Refresh
DataGrid1.Refresh

End Sub


Private Sub DataGrid1_Click()

DataGrid1.Col = 0
Mjeku.Text15.Text = Mjeku.DataGrid1.Text

DataGrid1.Col = 1
Mjeku.Text16.Text = Mjeku.DataGrid1.Text
Mjeku.Combo7.Text = Mjeku.DataGrid1.Text

DataGrid1.Col = 2
Mjeku.Text17.Text = Mjeku.DataGrid1.Text

DataGrid1.Col = 3
Mjeku.Text18.Text = Mjeku.DataGrid1.Text
Mjeku.Combo8.Text = Mjeku.DataGrid1.Text

DataGrid1.Col = 4
Mjeku.Combo9.Text = Mjeku.DataGrid1.Text

DataGrid1.Col = 5
Mjeku.DTPicker3 = Mjeku.DataGrid1.Text

DataGrid1.Col = 6
Mjeku.Combo10.Text = Mjeku.DataGrid1.Text

DataGrid1.Col = 7
Mjeku.Text19.Text = Mjeku.DataGrid1.Text

DataGrid1.Col = 8
Mjeku.Combo11.Text = Mjeku.DataGrid1.Text

DataGrid1.Col = 9
Mjeku.Combo12.Text = Mjeku.DataGrid1.Text

DataGrid1.Col = 10
Mjeku.Combo16.Text = Mjeku.DataGrid1.Text

Con = "Provider=MSDASQL.1;Data Source=Juli;Initial Catalog=SQLEXPRESS"
Con.Open "Driver=SQL Server;Server=.\SQLEXPRESS;Database=Hospital;uid=;pwd=;"
rec.Open "select IdTipInjeksioni from TipInjeksioni Where TipiInjeksionit = '" & Combo16 & "'", Con, adOpenUnspecified, adLockReadOnly
While Not rec.EOF
Text23.Text = rec(0)
rec.MoveNext
Wend
Con.Close


DataGrid1.Col = 11
Mjeku.Combo15.Text = Mjeku.DataGrid1.Text

Con = "Provider=MSDASQL.1;Data Source=Juli;Initial Catalog=SQLEXPRESS"
Con.Open "Driver=SQL Server;Server=.\SQLEXPRESS;Database=Hospital;uid=;pwd=;"
rec.Open "select IdInjeksion from Injeksionet Where EmerInjeksion = '" & Combo15 & "'", Con, adOpenUnspecified, adLockReadOnly
While Not rec.EOF
Text22.Text = rec(0)
rec.MoveNext
Wend
Con.Close

DataGrid1.Col = 12
Mjeku.Combo13.Text = Mjeku.DataGrid1.Text

Con = "Provider=MSDASQL.1;Data Source=Juli;Initial Catalog=SQLEXPRESS"
Con.Open "Driver=SQL Server;Server=.\SQLEXPRESS;Database=Hospital;uid=;pwd=;"
rec.Open "select IdPatologjia from Patologjia Where EmerPatologjia = '" & Combo13 & "'", Con, adOpenUnspecified, adLockReadOnly
While Not rec.EOF
Text20.Text = rec(0)
rec.MoveNext
Wend
Con.Close


DataGrid1.Col = 13
Mjeku.Combo14.Text = Mjeku.DataGrid1.Text

Con = "Provider=MSDASQL.1;Data Source=Juli;Initial Catalog=SQLEXPRESS"
Con.Open "Driver=SQL Server;Server=.\SQLEXPRESS;Database=Hospital;uid=;pwd=;"
rec.Open "select IdAnaliza from Analiza Where EmerAnaliza = '" & Combo14 & "'", Con, adOpenUnspecified, adLockReadOnly
While Not rec.EOF
Text21.Text = rec(0)
rec.MoveNext
Wend
Con.Close

End Sub
'Mbushja e combobox te Patologjise
Private Sub Combo13_Add()
Con = "Provider=MSDASQL.1;Data Source=Juli;Initial Catalog=SQLEXPRESS"
Con.Open "Driver=SQL Server;Server=.\SQLEXPRESS;Database=Hospital;uid=;pwd=;"
rec.Open "select EmerPatologjia from Patologjia", Con, adOpenUnspecified, adLockReadOnly
Combo13.AddItem "Selekto"
Combo13.Text = Me.Combo13.List(0)
While Not rec.EOF
Combo13.AddItem rec(0)
rec.MoveNext
Wend
Con.Close


End Sub

Private Sub Combo13_Click()
Con = "Provider=MSDASQL.1;Data Source=Juli;Initial Catalog=SQLEXPRESS"
Con.Open "Driver=SQL Server;Server=.\SQLEXPRESS;Database=Hospital;uid=;pwd=;"
rec.Open "select IdPatologjia,Pershkrimi from Patologjia Where EmerPatologjia = '" & Combo3 & "'", Con, adOpenUnspecified, adLockReadOnly
While Not rec.EOF
Text20.Text = rec(0)
rec.MoveNext
Wend
Con.Close
End Sub


'Mbushja e combobox te Analizes

Private Sub Combo14_Add()

Con = "Provider=MSDASQL.1;Data Source=Juli;Initial Catalog=SQLEXPRESS"
Con.Open "Driver=SQL Server;Server=.\SQLEXPRESS;Database=Hospital;uid=;pwd=;"
rec.Open "select EmerAnaliza from Analiza", Con, adOpenUnspecified, adLockReadOnly
Combo14.AddItem "Selekto"
Combo14.Text = Me.Combo14.List(0)
While Not rec.EOF
Combo14.AddItem rec(0)
rec.MoveNext
Wend
Con.Close

End Sub

Private Sub Combo14_Click()
Con = "Provider=MSDASQL.1;Data Source=Juli;Initial Catalog=SQLEXPRESS"
Con.Open "Driver=SQL Server;Server=.\SQLEXPRESS;Database=Hospital;uid=;pwd=;"
rec.Open "select IdAnaliza from Analiza Where EmerAnaliza = '" & Combo14 & "'", Con, adOpenUnspecified, adLockReadOnly
While Not rec.EOF
Text21.Text = rec(0)
rec.MoveNext
Wend
Con.Close
End Sub

'Mbushja e combobox te Injeksionit

Private Sub Combo15_Add()
Con = "Provider=MSDASQL.1;Data Source=Juli;Initial Catalog=SQLEXPRESS"
Con.Open "Driver=SQL Server;Server=.\SQLEXPRESS;Database=Hospital;uid=;pwd=;"
rec.Open "select EmerInjeksion from Injeksionet", Con, adOpenUnspecified, adLockReadOnly
Combo15.AddItem "Selekto"
Combo15.Text = Me.Combo15.List(0)
While Not rec.EOF
Combo15.AddItem rec(0)
rec.MoveNext
Wend
Con.Close
End Sub

Private Sub Combo15_Click()
Con = "Provider=MSDASQL.1;Data Source=Juli;Initial Catalog=SQLEXPRESS"
Con.Open "Driver=SQL Server;Server=.\SQLEXPRESS;Database=Hospital;uid=;pwd=;"
rec.Open "select IdInjeksion from Injeksionet Where EmerInjeksion = '" & Combo15 & "'", Con, adOpenUnspecified, adLockReadOnly
While Not rec.EOF
Text22.Text = rec(0)
rec.MoveNext
Wend
Con.Close
End Sub

'Mbushja e Combobox Tip Injeksioni

Private Sub Combo16_Add()

Con = "Provider=MSDASQL.1;Data Source=Juli;Initial Catalog=SQLEXPRESS"
Con.Open "Driver=SQL Server;Server=.\SQLEXPRESS;Database=Hospital;uid=;pwd=;"
rec.Open "select TipiInjeksionit from TipInjeksioni", Con, adOpenUnspecified, adLockReadOnly
Combo16.AddItem "Selekto"
Combo16.Text = Me.Combo16.List(0)
While Not rec.EOF
Combo16.AddItem rec(0)
rec.MoveNext
Wend
Con.Close

End Sub

Private Sub Combo16_Click()

Con = "Provider=MSDASQL.1;Data Source=Juli;Initial Catalog=SQLEXPRESS"
Con.Open "Driver=SQL Server;Server=.\SQLEXPRESS;Database=Hospital;uid=;pwd=;"
rec.Open "select IdTipInjeksioni from TipInjeksioni Where TipiInjeksionit = '" & Combo5 & "'", Con, adOpenUnspecified, adLockReadOnly
While Not rec.EOF
Text23.Text = rec(0)
rec.MoveNext
Wend
Con.Close

End Sub

'=============================================================== VIZITAT ===========================================================================================================
Private Sub DataGrid3_Click()
DataGrid3.Col = 0
Mjeku.Text32.Text = Mjeku.DataGrid3.Text

DataGrid3.Col = 1
Mjeku.Combo19.Text = Mjeku.DataGrid3.Text

DataGrid3.Col = 2
Mjeku.Text33.Text = Mjeku.DataGrid3.Text

DataGrid3.Col = 3
Mjeku.Combo20.Text = Mjeku.DataGrid3.Text


End Sub

Private Sub Combo19_Add()
Combo19.Clear
Dim strng As String
Con = "Provider=MSDASQL.1;Data Source=Juli;Initial Catalog=SQLEXPRESS"
Con.Open "Driver=SQL Server;Server=.\SQLEXPRESS;Database=Hospital;uid=;pwd=;"
rec.Open "select Emri from Pacienti", Con, adOpenUnspecified, adLockReadOnly

Combo19.AddItem "Selekto"
Combo19.Text = Me.Combo19.List(0)

While Not rec.EOF
Combo19.AddItem rec(0)
rec.MoveNext
Wend
Con.Close
End Sub

Private Sub Combo19_Click()
If Combo19.ListIndex = 0 Then
        Adodc3.RecordSource = "Select Distinct Pacienti.IdPacienti,Pacienti.Emri, Pacienti.Atesia, Pacienti.Mbiemri  From Pacienti"
    Else
        Adodc3.RecordSource = "Select Distinct Pacienti.IdPacienti,Pacienti.Emri, Pacienti.Atesia, Pacienti.Mbiemri From Pacienti Where Emri = '" & Combo19 & "'"
    End If
    
    Combo19.Refresh
    Adodc3.Refresh
DataGrid3.Refresh
Combo20.Enabled = True

Con = "Provider=MSDASQL.1;Data Source=Juli;Initial Catalog=SQLEXPRESS"
Con.Open "Driver=SQL Server;Server=.\SQLEXPRESS;Database=Hospital;uid=;pwd=;"
rec.Open "Select  Mbiemri from Pacienti Where Emri = '" & Combo19 & "'", Con, adOpenUnspecified, adLockReadOnly

Combo20.AddItem "Selekto"
Combo20.Text = Me.Combo20.List(0)

While Not rec.EOF
Combo20.AddItem rec(0)
rec.MoveNext
Wend
Con.Close

End Sub

Private Sub Combo20_Click()
If Combo20.ListIndex = 0 Then
        Adodc3.RecordSource = " Select Distinct Pacienti.IdPacienti,Pacienti.Emri, Pacienti.Atesia, Pacienti.Mbiemri, Pacienti.Gjinia, Pacienti.Datelindja, Pacienti.Vendlindja, Pacienti.Kontakt From Pacienti & "
    Else
               Adodc3.RecordSource = " Select Distinct Pacienti.IdPacienti,Pacienti.Emri, Pacienti.Atesia, Pacienti.Mbiemri From Pacienti Where Emri = '" & Combo19 & "' And Mbiemri = '" & Combo20 & "' "
    End If
    
    Combo20.Refresh
    Adodc3.Refresh
DataGrid3.Refresh

End Sub

'Klikimi i butonit I RI
Private Sub Command11_Click()
Text32.Text = " "
Text33.Text = " "
Combo19.Text = "Selekto"
Combo20.Text = "Selekto"
End Sub

'Klikimi i butonit RUAJ

Private Sub Command12_Click()
status = "Aktive"
Con = "Provider=MSDASQL.1;Data Source=Juli;Initial Catalog=SQLEXPRESS"
Con.Open "Driver=SQL Server;Server=.\SQLEXPRESS;Database=Hospital;uid=;pwd=;"

Set rs = Con.Execute( _
        "INSERT INTO Vizita (IdMjeku, IdPacienti, Status, DateVizita  ) VALUES('" & Text13.Text & "','" & Text32.Text & "', '" & status & "', '" & Format$(DTPicker5.Value, "yyyy/MM/dd HH:mm:ss") & "') ")
        
        
        

MsgBox "RUAJTJA E TE DHENAVE U KRYE ME SUKSES."

Con.Close
ProcExit:
Exit Sub
ProcError:

MsgBox "TE DHENAT NUK U RUAJTEN. SIGUROHUNI QE KENI PLOTESUAR TE GJITHA FUSHAT."
Con.Close
Resume ProcExit
Con.Close
Adodc4.RecordSource = "Select IdVizita, Pacienti.IdPacienti, Emri, Atesia, Mbiemri, DateVizita, Status From Vizita, Pacienti Where Vizita.Idpacienti = Pacienti.IdPacienti"
Adodc4.Refresh
DataGrid4.Refresh

End Sub

'Klikimi i butonit RIRUAJ
Private Sub Command13_Click()
Mjeku.SSTab3.Tab = 1
End Sub

'Klikimi i butonit MBYLL
Private Sub Command14_Click()
Unload Me
End Sub
'--------------------------- NDRYSHO PACIENT ---------------------------------------------------------------
Private Sub Combo21_Add()
Combo21.Clear
Dim strng As String
Con = "Provider=MSDASQL.1;Data Source=Juli;Initial Catalog=SQLEXPRESS"
Con.Open "Driver=SQL Server;Server=.\SQLEXPRESS;Database=Hospital;uid=;pwd=;"
rec.Open "select Emri from Pacienti", Con, adOpenUnspecified, adLockReadOnly

Combo21.AddItem "Selekto"
Combo21.Text = Me.Combo21.List(0)

While Not rec.EOF
Combo21.AddItem rec(0)
rec.MoveNext
Wend
Con.Close
End Sub

Private Sub Combo21_Click()
If Combo21.ListIndex = 0 Then
        Adodc4.RecordSource = "Select Distinct IdVizita, Pacienti.IdPacienti, Emri, Atesia, Mbiemri, DateVizita, Status From Vizita, Pacienti Where Vizita.Idpacienti = Pacienti.IdPacienti"
    Else
        Adodc4.RecordSource = "Select Distinct IdVizita, Pacienti.IdPacienti, Emri, Atesia, Mbiemri, DateVizita, Status From Vizita, Pacienti Where Vizita.Idpacienti = Pacienti.IdPacienti And Emri = '" & Combo21 & "'"
    End If
    
    Combo21.Refresh
    Adodc4.Refresh
DataGrid4.Refresh
Combo22.Enabled = True

Con = "Provider=MSDASQL.1;Data Source=Juli;Initial Catalog=SQLEXPRESS"
Con.Open "Driver=SQL Server;Server=.\SQLEXPRESS;Database=Hospital;uid=;pwd=;"
rec.Open "Select  Mbiemri from Pacienti Where Emri = '" & Combo21 & "'", Con, adOpenUnspecified, adLockReadOnly

Combo22.AddItem "Selekto"
Combo22.Text = Me.Combo22.List(0)

While Not rec.EOF
Combo22.AddItem rec(0)
rec.MoveNext
Wend
Con.Close
End Sub

Private Sub Combo22_Click()
If Combo22.ListIndex = 0 Then
        Adodc4.RecordSource = "Select Distinct IdVizita, Pacienti.IdPacienti, Emri, Atesia, Mbiemri, DateVizita, Status From Vizita, Pacienti Where Vizita.Idpacienti = Pacienti.IdPacienti"
    Else
               Adodc4.RecordSource = "Select Distinct IdVizita, Pacienti.IdPacienti, Emri, Atesia, Mbiemri, DateVizita, Status From Vizita, Pacienti Where Vizita.Idpacienti = Pacienti.IdPacienti And Emri = '" & Combo21 & "' And Mbiemri = '" & Combo22 & "' "
    End If
    
    Combo22.Refresh
    Adodc4.Refresh
DataGrid4.Refresh

End Sub
'Klikimi i butonit i ri
Private Sub Command16_Click()
Text37.Text = ""
Text35.Text = ""
Combo21.Text = ""
Text36.Text = ""
Combo22.Text = ""

End Sub

'Klikimi i butonit Riruaj
Private Sub Command17_Click()

Con = "Provider=MSDASQL.1;Data Source=Juli;Initial Catalog=SQLEXPRESS"
Con.Open "Driver=SQL Server;Server=.\SQLEXPRESS;Database=Hospital;uid=;pwd=;"

If Text37.Text = "" Or Text35.Text = "" Or Combo21.Text = "" Or Text36.Text = "" Or Combo22.Text = "" Then
MsgBox "PLOTESO FUSHAT."
Else
    On Error GoTo ProcError
    

Set rs = Con.Execute("Update Vizita Set  Vizita.IdMjeku= '" & Text13.Text & "', Vizita.IdPacienti = '" & Text35.Text & "', Vizita.DateVizita = '" & Format$(DTPicker6.Value, "yyyy.mm.dd HH:mm:ss") & "', Vizita.Status = '" & Combo26 & "' where Vizita.IdVizita = '" & Text37.Text & "' ")
 
MsgBox "MODIFIKIMI I TE DHENAVE U KRYE ME SUKSES."

End If
Con.Close
ProcExit:
Exit Sub
ProcError:


MsgBox "Error Number: " & Err.Number & " With The Description ->> " & Err.Description & " <<- Occured."

MsgBox "TE DHENAT NUK U MODIFIKUAN."
Con.Close
Resume ProcExit
Con.Close

End Sub

'Klikimi i butonit FSHIJ
Private Sub Command18_Click()
status = "E Anulluar"
Con = "Provider=MSDASQL.1;Data Source=Juli;Initial Catalog=SQLEXPRESS"
Con.Open "Driver=SQL Server;Server=.\SQLEXPRESS;Database=Hospital;uid=;pwd=;"
If Text37.Text = "" Then
MsgBox "PLOTESO FUSHAT."
Else
    On Error GoTo ProcError
    
Set rs = Con.Execute("Update Vizita Set Status = '" & status & "' Where IdVizita = '" & Text37.Text & "' ")
    
 
MsgBox "MODIFIKIMI I TE DHENAVE U KRYE ME SUKSES."

End If
Con.Close
ProcExit:
Exit Sub
ProcError:


MsgBox "Error Number: " & Err.Number & " With The Description ->> " & Err.Description & " <<- Occured."

MsgBox "TE DHENAT NUK U MODIFIKUAN."
Con.Close
Resume ProcExit
Con.Close

End Sub
'Klikimi i butonit REFRESH
Private Sub Command19_Click()
Text37.Text = ""
Text35.Text = ""
Combo21.Text = "Selekto"
Text36.Text = ""
Combo22.Text = "Selekto"
Combo26.Text = ""
DTPicker6.Value = Format(Now(), "dd/MM/yyyy")

Adodc4.RecordSource = "Select IdVizita, Pacienti.IdPacienti, Emri, Atesia, Mbiemri, DateVizita, Status From Vizita, Pacienti Where Vizita.Idpacienti = Pacienti.IdPacienti"
Adodc4.Refresh
DataGrid4.Refresh
End Sub

'Klikimi i butonit MBYLL
Private Sub Command20_Click()
Unload Me
End Sub


Private Sub DataGrid4_Click()
DataGrid4.Col = 0
Mjeku.Text37.Text = Mjeku.DataGrid4.Text

DataGrid4.Col = 1
Mjeku.Text35.Text = Mjeku.DataGrid4.Text

DataGrid4.Col = 2
Mjeku.Combo21.Text = Mjeku.DataGrid4.Text

DataGrid4.Col = 3
Mjeku.Text36.Text = Mjeku.DataGrid4.Text

DataGrid4.Col = 4
Mjeku.Combo22.Text = Mjeku.DataGrid4.Text


DataGrid4.Col = 5
Mjeku.DTPicker6.Value = Mjeku.DataGrid4.Text

DataGrid4.Col = 6
Mjeku.Combo26.Text = Mjeku.DataGrid4.Text


End Sub
'================================================== FUND SHTIMI I VIZITAVE =====================================================================================


'================================================== GJENERIMI I FATURES ====================================================================================

Private Sub Combo17_Add()
Combo17.Clear
Dim strng As String
Con = "Provider=MSDASQL.1;Data Source=Juli;Initial Catalog=SQLEXPRESS"
Con.Open "Driver=SQL Server;Server=.\SQLEXPRESS;Database=Hospital;uid=;pwd=;"
rec.Open "select Emri from Pacienti", Con, adOpenUnspecified, adLockReadOnly

Combo17.AddItem "Selekto"
Combo17.Text = Me.Combo17.List(0)

While Not rec.EOF
Combo17.AddItem rec(0)
rec.MoveNext
Wend
Con.Close
End Sub

Private Sub Combo17_Click()
If Combo17.ListIndex = 0 Then
        Adodc2.RecordSource = "Select Distinct Pacienti.IdPacienti,Pacienti.Emri, Pacienti.Atesia, Pacienti.Mbiemri, Pacienti.Gjinia, Pacienti.Datelindja, Pacienti.Vendlindja, Pacienti.Kontakt From Pacienti"
    Else
        Adodc2.RecordSource = "Select Distinct Pacienti.IdPacienti,Pacienti.Emri, Pacienti.Atesia, Pacienti.Mbiemri, Pacienti.Gjinia, Pacienti.Datelindja, Pacienti.Vendlindja, Pacienti.Kontakt From Pacienti Where Emri = '" & Combo17 & "'"
    End If
    
    Combo17.Refresh
    Adodc2.Refresh
DataGrid2.Refresh
Combo18.Enabled = True

Con = "Provider=MSDASQL.1;Data Source=Juli;Initial Catalog=SQLEXPRESS"
Con.Open "Driver=SQL Server;Server=.\SQLEXPRESS;Database=Hospital;uid=;pwd=;"
rec.Open "Select Distinct Mbiemri from Pacienti Where Emri = '" & Combo17 & "'", Con, adOpenUnspecified, adLockReadOnly

Combo18.AddItem "Selekto"
Combo18.Text = Me.Combo18.List(0)

While Not rec.EOF
Combo18.AddItem rec(0)
rec.MoveNext
Wend
Con.Close

End Sub

Private Sub Combo18_Click()
If Combo18.ListIndex = 0 Then
        Adodc2.RecordSource = " Select Distinct Pacienti.IdPacienti,Pacienti.Emri, Pacienti.Atesia, Pacienti.Mbiemri, Pacienti.Gjinia, Pacienti.Datelindja, Pacienti.Vendlindja, Pacienti.Kontakt From Pacienti & "
    Else
               Adodc2.RecordSource = " Select Distinct Pacienti.IdPacienti,Pacienti.Emri, Pacienti.Atesia, Pacienti.Mbiemri, Pacienti.Gjinia, Pacienti.Datelindja, Pacienti.Vendlindja, Pacienti.Kontakt From Pacienti Where Emri = '" & Combo17 & "' And Mbiemri = '" & Combo18 & "' "
    End If
    
    Combo18.Refresh
    Adodc2.Refresh
DataGrid2.Refresh

End Sub

Private Sub DataGrid2_Click()
DataGrid2.Col = 0
Mjeku.Text24.Text = Mjeku.DataGrid2.Text

DataGrid2.Col = 1
Mjeku.Text25.Text = Mjeku.DataGrid2.Text
Mjeku.Combo17.Text = Mjeku.DataGrid2.Text

DataGrid2.Col = 2
Mjeku.Text26.Text = Mjeku.DataGrid2.Text

DataGrid2.Col = 3
Mjeku.Text27.Text = Mjeku.DataGrid2.Text
Mjeku.Combo18.Text = Mjeku.DataGrid2.Text

DataGrid2.Col = 4
Mjeku.Text28.Text = Mjeku.DataGrid2.Text

DataGrid2.Col = 5
Mjeku.DTPicker4 = Mjeku.DataGrid2.Text

End Sub


Private Sub Command10_Click()
Dim a As Integer
Text31.Text = ""
For a = 0 To List1.ListCount - 1
Text31.Text = Text31.Text & List1.List(a) & vbNewLine
Next a


Dim b As Integer
Text14.Text = ""
For b = 0 To List2.ListCount - 1
Text14.Text = Text14.Text & List2.List(b) & vbNewLine
Next b

Dim mosha As Integer

If Text24.Text = "" Or Text25.Text = "" Or Text26.Text = "" Or Text27.Text = "" Or Text28.Text = "" Then
MsgBox "Ju lutem zgjidhni nje pacient"
Else
mosha = DTPicker2.Value - DTPicker4.Value
mosha = mosha / 365

Text29.Text = DTPicker4.Value
Text30.Text = DTPicker2.Value
Label37.Caption = " ================ FATURE ================                                                                Data: " & Format$(DTPicker2.Value, "dd/MM/yyyy") & ""

Label47.Caption = " " & Text24.Text & ""
Label52.Caption = " " & Text25.Text & ""
Label58.Caption = " " & Text26.Text & ""
Label59.Caption = " " & Text27.Text & ""
Label49.Caption = " " & Text28.Text & ""
Label55.Caption = " " & Text29.Text & ""
Label57.Caption = " " & mosha & ""
End If

Dim tot As Integer
For i = 0 To List2.ListCount - 1
tot = tot + Val(List2.List(i))
Next

Label46.Caption = tot
End Sub

Private Sub Combo23_Add()
Con = "Provider=MSDASQL.1;Data Source=Juli;Initial Catalog=SQLEXPRESS"
Con.Open "Driver=SQL Server;Server=.\SQLEXPRESS;Database=Hospital;uid=;pwd=;"
rec.Open "select EmerIlace from Ilace", Con, adOpenUnspecified, adLockReadOnly
Combo23.AddItem "Selekto"
Combo23.Text = Me.Combo23.List(0)
While Not rec.EOF
Combo23.AddItem rec(0)
rec.MoveNext
Wend
Con.Close

End Sub

Private Sub Combo23_Click()
Con = "Provider=MSDASQL.1;Data Source=Juli;Initial Catalog=SQLEXPRESS"
Con.Open "Driver=SQL Server;Server=.\SQLEXPRESS;Database=Hospital;uid=;pwd=;"
rec.Open "select EmerIlace,CmimiIlace from Ilace Where EmerIlace = '" & Combo23 & "'", Con, adOpenUnspecified, adLockReadOnly
While Not rec.EOF
List1.Text = rec(0)
Text38.Text = rec(1)
rec.MoveNext
Wend
Con.Close
End Sub
'
Private Sub Combo24_Add()
Con = "Provider=MSDASQL.1;Data Source=Juli;Initial Catalog=SQLEXPRESS"
Con.Open "Driver=SQL Server;Server=.\SQLEXPRESS;Database=Hospital;uid=;pwd=;"
rec.Open "select EmerAnaliza from Analiza", Con, adOpenUnspecified, adLockReadOnly
Combo24.AddItem "Selekto"
Combo24.Text = Me.Combo24.List(0)
While Not rec.EOF
Combo24.AddItem rec(0)
rec.MoveNext
Wend
Con.Close

End Sub

Private Sub Combo24_Click()
Con = "Provider=MSDASQL.1;Data Source=Juli;Initial Catalog=SQLEXPRESS"
Con.Open "Driver=SQL Server;Server=.\SQLEXPRESS;Database=Hospital;uid=;pwd=;"
rec.Open "select EmerAnaliza,Kosto from Analiza Where EmerAnaliza = '" & Combo24 & "'", Con, adOpenUnspecified, adLockReadOnly
While Not rec.EOF
List1.Text = rec(0)
Text39.Text = rec(1)
rec.MoveNext
Wend
Con.Close
End Sub

Private Sub Combo25_Add()
Con = "Provider=MSDASQL.1;Data Source=Juli;Initial Catalog=SQLEXPRESS"
Con.Open "Driver=SQL Server;Server=.\SQLEXPRESS;Database=Hospital;uid=;pwd=;"
rec.Open "select EmerInjeksion from Injeksionet", Con, adOpenUnspecified, adLockReadOnly
Combo25.AddItem "Selekto"
Combo25.Text = Me.Combo25.List(0)
While Not rec.EOF
Combo25.AddItem rec(0)
rec.MoveNext
Wend
Con.Close

End Sub

Private Sub Combo25_Click()
Con = "Provider=MSDASQL.1;Data Source=Juli;Initial Catalog=SQLEXPRESS"
Con.Open "Driver=SQL Server;Server=.\SQLEXPRESS;Database=Hospital;uid=;pwd=;"
rec.Open "select EmerInjeksion,CmimInjeksion from Injeksionet  Where EmerInjeksion = '" & Combo25 & "'", Con, adOpenUnspecified, adLockReadOnly
While Not rec.EOF
List1.Text = rec(0)
Text40.Text = rec(1)
rec.MoveNext
Wend
Con.Close
End Sub

Private Sub Command24_Click()
If Combo23.Visible = False And Combo24.Visible = False And Combo25.Visible = False Then
MsgBox ("Ju lutem zgjidhni nje ILAC,ANALIZE Ose INJEKSION te marre")

ElseIf Combo23.Visible = True Then
List1.AddItem Combo23.Text
List2.AddItem Text38.Text

ElseIf Combo24.Visible = True Then

List1.AddItem Combo24.Text
List2.AddItem Text39.Text

ElseIf Combo25.Visible = True Then
List1.AddItem Combo25.Text
List2.AddItem Text40.Text

End If


End Sub

Private Sub Command25_Click()
Text24.Text = ""
Label47.Caption = ""
Text25.Text = ""
Label52.Caption = ""
Text26.Text = ""
Label58.Caption = ""
Text27.Text = ""
Label59.Caption = ""
Text28.Text = ""
Combo17.Text = "Selekto"
Combo18.Text = ""
Combo18.Enabled = False
Label49.Caption = ""
Label55.Caption = ""
Label57.Caption = ""
List1.Clear
List2.Clear
Text31.Text = ""
Text14.Text = ""

End Sub
'Klikimi i butonit dhe shfaqja e ComboBox.
Private Sub Command21_Click()
Combo23.Visible = True
Combo24.Visible = False
Combo25.Visible = False
End Sub

Private Sub Command22_Click()
Combo23.Visible = False
Combo24.Visible = True
Combo25.Visible = False
End Sub

Private Sub Command23_Click()
Combo23.Visible = False
Combo24.Visible = False
Combo25.Visible = True
End Sub

Private Sub Command26_Click()

Dim ind As Integer
Dim ind2 As Integer
Dim i, j As Integer
'Dim a, b As Integer

ind = List1.ListIndex
ind2 = List2.ListIndex

If List1.Text = "" And List2.Text = "" Then
MsgBox (" Nuk ka asnje produkt te zgjedhur !! ")
Else
If List1.Selected(i) <> List2.Selected(j) Then
MsgBox ("Ju lutem zgjidhni produktin me cmimin perkates !! ")
Else
If ind >= 0 And ind2 >= 0 Then
List1.RemoveItem ind
List2.RemoveItem ind2
End If
End If
End If
End Sub



Private Sub Command27_Click()


Con = "Provider=MSDASQL.1;Data Source=Juli;Initial Catalog=SQLEXPRESS"
Con.Open "Driver=SQL Server;Server=.\SQLEXPRESS;Database=Hospital;uid=;pwd=;"
If Text24.Text = "" Or Text25.Text = "" Or Text26.Text = "" Or Text27.Text = "" Or Text28 = "" Then
MsgBox "PLOTESONI TE GJITHA FUSHAT"

Else
    On Error GoTo ProcError
Set rs = Con.Execute( _
        "INSERT INTO Fatura(IdPacient, IdDoktor, KostoTotale, Data )VALUES('" & Text24.Text & "','" & Text13.Text & "','" & Label46.Caption & "','" & Format$(DTPicker2.Value, "yyyy.mm.dd") & "')")
MsgBox "RUAJTJA E TE DHENAVE U KRYE ME SUKSES."
Dim i As Integer
Static wd1 As Word.Application
Static wd1Doc As Word.Document
Set wd1 = New Word.Application

wd1.Visible = True
Set wd1Doc = wd1.Documents.Add(App.Path & "\fature.Dotx")


With wdDoc
wd1Doc.Bookmarks("Data").Range.Text = Text30.Text
wd1Doc.Bookmarks("Id").Range.Text = Label47.Caption
wd1Doc.Bookmarks("Emri").Range.Text = Label52.Caption
wd1Doc.Bookmarks("Atesia").Range.Text = Label58.Caption
wd1Doc.Bookmarks("Mbiemri").Range.Text = Label59.Caption
wd1Doc.Bookmarks("Gjinia").Range.Text = Label49.Caption
wd1Doc.Bookmarks("Datelindja").Range.Text = Label55.Caption
wd1Doc.Bookmarks("Mosha").Range.Text = Label57.Caption

wd1Doc.Bookmarks("Produktet").Range.Text = Text31.Text
wd1Doc.Bookmarks("Cmimet").Range.Text = Text14.Text
wd1Doc.Bookmarks("Total").Range.Text = Label46.Caption

End With
Set wd1 = Nothing
Set wd1Doc = Nothing



End If
Con.Close
ProcExit:
Exit Sub
ProcError:

MsgBox "TE DHENAT NUK U RUAJTEN. SIGUROHUNI QE KENI PLOTESUAR TE GJITHA FUSHAT."
Con.Close
Resume ProcExit
Con.Close



End Sub

Private Sub Command28_Click()
PacientRaport.Show
End Sub

Private Sub Command29_Click()
FatureRaport.Show
End Sub

'============================================================= FUND FATURA =============================================================================

'LOGOUT
Private Sub Command15_Click()
Unload Me
Login.Show
Login.Text1.Text = ""
Login.Text2.Text = ""
Login.Text1.SetFocus
End Sub

Private Sub Form_Load()
Combo3_Add
Combo4_Add
Combo5_Add
Combo6_Add
Combo7_Add
Combo13_Add
Combo14_Add
Combo15_Add
Combo16_Add
Combo17_Add
Combo19_Add
Combo21_Add
Combo23_Add
Combo24_Add
Combo25_Add
DTPicker1.Value = Format(Now(), "dd/MM/yyyy")
DTPicker2.Value = Format(Now(), "dd/MM/yyyy")
DTPicker3.Value = Format(Now(), "dd/MM/yyyy")
DTPicker5.Value = Format(Now(), "dd/MM/yyyy")
DTPicker6.Value = Format(Now(), "dd/MM/yyyy")
End Sub





'MENUTE

Private Sub shtoPacient_Click()
Mjeku.SSTab1.Tab = 0
Mjeku.SSTab2.Tab = 0
End Sub
Private Sub ndryshoPacient_Click()
Mjeku.SSTab1.Tab = 0
Mjeku.SSTab2.Tab = 1
End Sub
Private Sub raportPacient_Click()
Mjeku.SSTab1.Tab = 3
End Sub

Private Sub shtoVizita_Click()
Mjeku.SSTab1.Tab = 1
Mjeku.SSTab3.Tab = 0
End Sub

Private Sub vizitateMia_Click()
Mjeku.SSTab1.Tab = 1
Mjeku.SSTab3.Tab = 1
End Sub
Private Sub fatura_Click()
Mjeku.SSTab1.Tab = 2
End Sub
Private Sub raport_Click()
Mjeku.SSTab1.Tab = 3
End Sub
