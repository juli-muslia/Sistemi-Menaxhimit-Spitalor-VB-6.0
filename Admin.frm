VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Admin 
   BackColor       =   &H00808000&
   Caption         =   "Hospital Management System - Admin"
   ClientHeight    =   11610
   ClientLeft      =   525
   ClientTop       =   2145
   ClientWidth     =   21360
   LinkTopic       =   "Form1"
   ScaleHeight     =   11610
   ScaleWidth      =   21360
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command48 
      Caption         =   "DIL"
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
      Picture         =   "Admin.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   191
      Top             =   0
      Width           =   1335
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   10215
      Left            =   480
      TabIndex        =   0
      Top             =   1200
      Width           =   20535
      _ExtentX        =   36221
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
      TabCaption(0)   =   "Mjeku"
      TabPicture(0)   =   "Admin.frx":4321
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "SSTab2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Departamenti"
      TabPicture(1)   =   "Admin.frx":433D
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "SSTab3"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Ilaçe / Analiza / Injeksione"
      TabPicture(2)   =   "Admin.frx":4359
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "SSTab4"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Raporte"
      TabPicture(3)   =   "Admin.frx":4375
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Command45"
      Tab(3).Control(1)=   "Command46"
      Tab(3).Control(2)=   "Command47"
      Tab(3).Control(3)=   "Command44"
      Tab(3).ControlCount=   4
      Begin VB.CommandButton Command44 
         Caption         =   "Gjenero Raport Mjeku"
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
         Left            =   -69480
         Picture         =   "Admin.frx":4391
         Style           =   1  'Graphical
         TabIndex        =   190
         Top             =   3120
         Width           =   2415
      End
      Begin VB.CommandButton Command47 
         Caption         =   "Gjenero Raport Injeksione"
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
         Left            =   -62280
         Picture         =   "Admin.frx":997B
         Style           =   1  'Graphical
         TabIndex        =   189
         Top             =   3120
         Width           =   2415
      End
      Begin VB.CommandButton Command46 
         Caption         =   "Gjenero Raport Analiza"
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
         Left            =   -64680
         Picture         =   "Admin.frx":EF65
         Style           =   1  'Graphical
         TabIndex        =   188
         Top             =   3120
         Width           =   2415
      End
      Begin VB.CommandButton Command45 
         Caption         =   "Gjenero Raport Ilace"
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
         Left            =   -67080
         Picture         =   "Admin.frx":1454F
         Style           =   1  'Graphical
         TabIndex        =   187
         Top             =   3120
         Width           =   2415
      End
      Begin TabDlg.SSTab SSTab4 
         Height          =   8295
         Left            =   -74520
         TabIndex        =   96
         Top             =   1440
         Width           =   19215
         _ExtentX        =   33893
         _ExtentY        =   14631
         _Version        =   393216
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
         TabCaption(0)   =   "Ilaçe"
         TabPicture(0)   =   "Admin.frx":19B39
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "SSTab5"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Analiza"
         TabPicture(1)   =   "Admin.frx":19B55
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "SSTab6"
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "Injeksione"
         TabPicture(2)   =   "Admin.frx":19B71
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "SSTab7"
         Tab(2).ControlCount=   1
         Begin TabDlg.SSTab SSTab7 
            Height          =   6315
            Left            =   -73080
            TabIndex        =   150
            Top             =   1080
            Width           =   14415
            _ExtentX        =   25426
            _ExtentY        =   11139
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
            TabCaption(0)   =   "Shto Injeksion"
            TabPicture(0)   =   "Admin.frx":19B8D
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "Label45"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).Control(1)=   "Label49"
            Tab(0).Control(1).Enabled=   0   'False
            Tab(0).Control(2)=   "Label50"
            Tab(0).Control(2).Enabled=   0   'False
            Tab(0).Control(3)=   "Label57"
            Tab(0).Control(3).Enabled=   0   'False
            Tab(0).Control(4)=   "Command35"
            Tab(0).Control(4).Enabled=   0   'False
            Tab(0).Control(5)=   "Command36"
            Tab(0).Control(5).Enabled=   0   'False
            Tab(0).Control(6)=   "Command37"
            Tab(0).Control(6).Enabled=   0   'False
            Tab(0).Control(7)=   "Command38"
            Tab(0).Control(7).Enabled=   0   'False
            Tab(0).Control(8)=   "Text39"
            Tab(0).Control(8).Enabled=   0   'False
            Tab(0).Control(9)=   "Text40"
            Tab(0).Control(9).Enabled=   0   'False
            Tab(0).Control(10)=   "Text41"
            Tab(0).Control(10).Enabled=   0   'False
            Tab(0).Control(11)=   "Combo17"
            Tab(0).Control(11).Enabled=   0   'False
            Tab(0).Control(12)=   "Text46"
            Tab(0).Control(12).Enabled=   0   'False
            Tab(0).ControlCount=   13
            TabCaption(1)   =   "Ndrysho Injeksion"
            TabPicture(1)   =   "Admin.frx":19BA9
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "DataGrid5"
            Tab(1).Control(1)=   "Command39"
            Tab(1).Control(2)=   "Command40"
            Tab(1).Control(3)=   "Command41"
            Tab(1).Control(4)=   "Command42"
            Tab(1).Control(5)=   "Command43"
            Tab(1).Control(6)=   "Frame9"
            Tab(1).Control(7)=   "Frame10"
            Tab(1).Control(8)=   "Adodc5"
            Tab(1).ControlCount=   9
            Begin VB.TextBox Text46 
               Height          =   375
               Left            =   8880
               TabIndex        =   184
               Text            =   "Text46"
               Top             =   3660
               Visible         =   0   'False
               Width           =   1095
            End
            Begin VB.ComboBox Combo17 
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
               Left            =   5760
               TabIndex        =   183
               Top             =   3660
               Width           =   2775
            End
            Begin MSAdodcLib.Adodc Adodc5 
               Height          =   495
               Left            =   -74280
               Top             =   540
               Visible         =   0   'False
               Width           =   2655
               _ExtentX        =   4683
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
               Connect         =   $"Admin.frx":19BC5
               OLEDBString     =   $"Admin.frx":19C64
               OLEDBFile       =   ""
               DataSourceName  =   ""
               OtherAttributes =   ""
               UserName        =   ""
               Password        =   ""
               RecordSource    =   $"Admin.frx":19D03
               Caption         =   "Adodc5"
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
            Begin VB.Frame Frame10 
               Caption         =   "Ndrysho Injeksion"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   4215
               Left            =   -66240
               TabIndex        =   167
               Top             =   1380
               Width           =   5415
               Begin VB.TextBox Text45 
                  Height          =   375
                  Left            =   240
                  TabIndex        =   181
                  Text            =   "Text45"
                  Top             =   2880
                  Visible         =   0   'False
                  Width           =   975
               End
               Begin VB.ComboBox Combo16 
                  Height          =   315
                  ItemData        =   "Admin.frx":19DAA
                  Left            =   2520
                  List            =   "Admin.frx":19DB4
                  TabIndex        =   180
                  Top             =   3000
                  Width           =   2535
               End
               Begin VB.ComboBox Combo15 
                  Height          =   315
                  Left            =   2520
                  TabIndex        =   179
                  Top             =   2520
                  Width           =   2535
               End
               Begin VB.TextBox Text44 
                  Height          =   375
                  Left            =   2520
                  TabIndex        =   178
                  Top             =   2040
                  Width           =   2535
               End
               Begin VB.TextBox Text43 
                  Height          =   375
                  Left            =   2520
                  TabIndex        =   177
                  Top             =   1560
                  Width           =   2535
               End
               Begin VB.TextBox Text42 
                  Enabled         =   0   'False
                  Height          =   375
                  Left            =   2520
                  TabIndex        =   176
                  Top             =   1080
                  Width           =   2535
               End
               Begin VB.Label Label56 
                  Caption         =   "Gjendje"
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
                  Left            =   1320
                  TabIndex        =   175
                  Top             =   3000
                  Width           =   1095
               End
               Begin VB.Label Label55 
                  Caption         =   "Tip Injeksioni"
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
                  Left            =   525
                  TabIndex        =   174
                  Top             =   2520
                  Width           =   1815
               End
               Begin VB.Label Label54 
                  Caption         =   "Çmim Injeksioni"
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
                  Left            =   150
                  TabIndex        =   173
                  Top             =   2040
                  Width           =   2175
               End
               Begin VB.Label Label53 
                  Caption         =   "Emer Injeksioni"
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
                  TabIndex        =   172
                  Top             =   1560
                  Width           =   2055
               End
               Begin VB.Label Label52 
                  Caption         =   "Kod Injeksioni"
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
                  TabIndex        =   171
                  Top             =   1080
                  Width           =   1935
               End
            End
            Begin VB.Frame Frame9 
               Caption         =   "Kerko Injeksion"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1815
               Left            =   -74640
               TabIndex        =   166
               Top             =   1380
               Width           =   8295
               Begin VB.ComboBox Combo14 
                  Height          =   315
                  Left            =   3000
                  TabIndex        =   170
                  Top             =   840
                  Width           =   2175
               End
               Begin VB.Label Label51 
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
                  Height          =   615
                  Left            =   480
                  TabIndex        =   169
                  Top             =   840
                  Width           =   2295
               End
            End
            Begin VB.CommandButton Command43 
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
               Left            =   -66360
               Picture         =   "Admin.frx":19DD1
               Style           =   1  'Graphical
               TabIndex        =   165
               Top             =   420
               Width           =   855
            End
            Begin VB.CommandButton Command42 
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
               Left            =   -67320
               Picture         =   "Admin.frx":1D4B3
               Style           =   1  'Graphical
               TabIndex        =   164
               Top             =   420
               Width           =   855
            End
            Begin VB.CommandButton Command41 
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
               Left            =   -68280
               Picture         =   "Admin.frx":208F4
               Style           =   1  'Graphical
               TabIndex        =   163
               Top             =   420
               Width           =   855
            End
            Begin VB.CommandButton Command40 
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
               Left            =   -69240
               Picture         =   "Admin.frx":23BC7
               Style           =   1  'Graphical
               TabIndex        =   162
               Top             =   420
               Width           =   855
            End
            Begin VB.CommandButton Command39 
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
               Left            =   -70200
               Picture         =   "Admin.frx":26D8E
               Style           =   1  'Graphical
               TabIndex        =   161
               Top             =   420
               Width           =   855
            End
            Begin VB.TextBox Text41 
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
               Left            =   5760
               TabIndex        =   160
               Top             =   3180
               Width           =   2775
            End
            Begin VB.TextBox Text40 
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
               Left            =   5760
               TabIndex        =   159
               Top             =   2700
               Width           =   2775
            End
            Begin VB.TextBox Text39 
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
               Left            =   5760
               TabIndex        =   155
               Top             =   2220
               Width           =   2775
            End
            Begin VB.CommandButton Command38 
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
               Left            =   7680
               Picture         =   "Admin.frx":2A079
               Style           =   1  'Graphical
               TabIndex        =   154
               Top             =   420
               Width           =   855
            End
            Begin VB.CommandButton Command37 
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
               Left            =   6720
               Picture         =   "Admin.frx":2D75B
               Style           =   1  'Graphical
               TabIndex        =   153
               Top             =   420
               Width           =   855
            End
            Begin VB.CommandButton Command36 
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
               Left            =   5760
               Picture         =   "Admin.frx":30922
               Style           =   1  'Graphical
               TabIndex        =   152
               Top             =   420
               Width           =   855
            End
            Begin VB.CommandButton Command35 
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
               Left            =   4800
               Picture         =   "Admin.frx":3399A
               Style           =   1  'Graphical
               TabIndex        =   151
               Top             =   420
               Width           =   855
            End
            Begin MSDataGridLib.DataGrid DataGrid5 
               Bindings        =   "Admin.frx":36C85
               Height          =   2175
               Left            =   -74880
               TabIndex        =   168
               Top             =   3420
               Width           =   8535
               _ExtentX        =   15055
               _ExtentY        =   3836
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
               ColumnCount     =   5
               BeginProperty Column00 
                  DataField       =   "IdInjeksion"
                  Caption         =   "IdInjeksion"
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
               BeginProperty Column02 
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
               BeginProperty Column03 
                  DataField       =   "CmimInjeksion"
                  Caption         =   "CmimInjeksion"
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
                  DataField       =   "Gjendje"
                  Caption         =   "Gjendje"
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
                     ColumnWidth     =   1739.906
                  EndProperty
               EndProperty
            End
            Begin VB.Label Label57 
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
               Left            =   3300
               TabIndex        =   182
               Top             =   3660
               Width           =   2160
            End
            Begin VB.Label Label50 
               Caption         =   "Çmimi i Injeksionit"
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
               Left            =   2900
               TabIndex        =   158
               Top             =   3180
               Width           =   2535
            End
            Begin VB.Label Label49 
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
               Left            =   3120
               TabIndex        =   157
               Top             =   2700
               Width           =   2295
            End
            Begin VB.Label Label45 
               Caption         =   "Kodi i Injeksionit"
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
               Left            =   3120
               TabIndex        =   156
               Top             =   2220
               Width           =   2295
            End
         End
         Begin TabDlg.SSTab SSTab6 
            Height          =   6315
            Left            =   -73080
            TabIndex        =   126
            Top             =   1080
            Width           =   14415
            _ExtentX        =   25426
            _ExtentY        =   11139
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
            TabCaption(0)   =   "Shto Analiza"
            TabPicture(0)   =   "Admin.frx":36C9A
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "Label42"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).Control(1)=   "Label43"
            Tab(0).Control(1).Enabled=   0   'False
            Tab(0).Control(2)=   "Command27"
            Tab(0).Control(2).Enabled=   0   'False
            Tab(0).Control(3)=   "Command28"
            Tab(0).Control(3).Enabled=   0   'False
            Tab(0).Control(4)=   "Command29"
            Tab(0).Control(4).Enabled=   0   'False
            Tab(0).Control(5)=   "Command30"
            Tab(0).Control(5).Enabled=   0   'False
            Tab(0).Control(6)=   "Text34"
            Tab(0).Control(6).Enabled=   0   'False
            Tab(0).Control(7)=   "Text35"
            Tab(0).Control(7).Enabled=   0   'False
            Tab(0).ControlCount=   8
            TabCaption(1)   =   "Ndrysho Analiza"
            TabPicture(1)   =   "Admin.frx":36CB6
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "Command31"
            Tab(1).Control(1)=   "Command32"
            Tab(1).Control(2)=   "Command33"
            Tab(1).Control(3)=   "Command34"
            Tab(1).Control(4)=   "Frame7"
            Tab(1).Control(5)=   "Frame8"
            Tab(1).Control(6)=   "Adodc4"
            Tab(1).ControlCount=   7
            Begin MSAdodcLib.Adodc Adodc4 
               Height          =   495
               Left            =   -74640
               Top             =   300
               Visible         =   0   'False
               Width           =   2895
               _ExtentX        =   5106
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
               Connect         =   $"Admin.frx":36CD2
               OLEDBString     =   $"Admin.frx":36D71
               OLEDBFile       =   ""
               DataSourceName  =   ""
               OtherAttributes =   ""
               UserName        =   ""
               Password        =   ""
               RecordSource    =   "Select *  from Analiza"
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
            Begin VB.Frame Frame8 
               Caption         =   "Ndrysho Analize"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   4215
               Left            =   -66840
               TabIndex        =   140
               Top             =   1380
               Width           =   5775
               Begin VB.TextBox Text38 
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
                  Left            =   2640
                  TabIndex        =   149
                  Top             =   1920
                  Width           =   2415
               End
               Begin VB.TextBox Text37 
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
                  Left            =   2640
                  TabIndex        =   148
                  Top             =   1440
                  Width           =   2415
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
                  Height          =   375
                  Left            =   2640
                  TabIndex        =   147
                  Top             =   960
                  Width           =   2415
               End
               Begin VB.Label Label48 
                  Caption         =   "Çmimi i Analizes"
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
                  Left            =   240
                  TabIndex        =   146
                  Top             =   1920
                  Width           =   2295
               End
               Begin VB.Label Label47 
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
                  TabIndex        =   145
                  Top             =   1440
                  Width           =   1935
               End
               Begin VB.Label Label46 
                  Caption         =   "Kodi i Analizes"
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
                  TabIndex        =   144
                  Top             =   960
                  Width           =   1935
               End
            End
            Begin VB.Frame Frame7 
               Caption         =   "Kerko Analize"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   4215
               Left            =   -74640
               TabIndex        =   139
               Top             =   1380
               Width           =   7575
               Begin MSDataGridLib.DataGrid DataGrid4 
                  Bindings        =   "Admin.frx":36E10
                  Height          =   1935
                  Left            =   240
                  TabIndex        =   143
                  Top             =   1920
                  Width           =   7095
                  _ExtentX        =   12515
                  _ExtentY        =   3413
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
                  ColumnCount     =   3
                  BeginProperty Column00 
                     DataField       =   "IdAnaliza"
                     Caption         =   "IdAnaliza"
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
                  BeginProperty Column02 
                     DataField       =   "Kosto"
                     Caption         =   "Kosto"
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
                  EndProperty
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
                  Left            =   2520
                  TabIndex        =   142
                  Top             =   960
                  Width           =   3615
               End
               Begin VB.Label Label44 
                  Caption         =   "Emer Analize"
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
                  TabIndex        =   141
                  Top             =   960
                  Width           =   1815
               End
            End
            Begin VB.CommandButton Command34 
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
               Left            =   -67320
               Picture         =   "Admin.frx":36E25
               Style           =   1  'Graphical
               TabIndex        =   138
               Top             =   420
               Width           =   855
            End
            Begin VB.CommandButton Command33 
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
               Left            =   -68280
               Picture         =   "Admin.frx":3A507
               Style           =   1  'Graphical
               TabIndex        =   137
               Top             =   420
               Width           =   855
            End
            Begin VB.CommandButton Command32 
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
               Left            =   -69240
               Picture         =   "Admin.frx":3D948
               Style           =   1  'Graphical
               TabIndex        =   136
               Top             =   420
               Width           =   855
            End
            Begin VB.CommandButton Command31 
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
               Left            =   -70200
               Picture         =   "Admin.frx":40B0F
               Style           =   1  'Graphical
               TabIndex        =   135
               Top             =   420
               Width           =   855
            End
            Begin VB.TextBox Text35 
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
               Left            =   5760
               TabIndex        =   134
               Top             =   3060
               Width           =   2775
            End
            Begin VB.TextBox Text34 
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
               Left            =   5760
               TabIndex        =   133
               Top             =   2580
               Width           =   2775
            End
            Begin VB.CommandButton Command30 
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
               Left            =   8400
               Picture         =   "Admin.frx":43DFA
               Style           =   1  'Graphical
               TabIndex        =   130
               Top             =   420
               Width           =   855
            End
            Begin VB.CommandButton Command29 
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
               Left            =   7440
               Picture         =   "Admin.frx":474DC
               Style           =   1  'Graphical
               TabIndex        =   129
               Top             =   420
               Width           =   855
            End
            Begin VB.CommandButton Command28 
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
               Left            =   6480
               Picture         =   "Admin.frx":4A6A3
               Style           =   1  'Graphical
               TabIndex        =   128
               Top             =   420
               Width           =   855
            End
            Begin VB.CommandButton Command27 
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
               Left            =   5520
               Picture         =   "Admin.frx":4D71B
               Style           =   1  'Graphical
               TabIndex        =   127
               Top             =   420
               Width           =   855
            End
            Begin VB.Label Label43 
               Caption         =   "Çmimi i Analizes"
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
               Left            =   3240
               TabIndex        =   132
               Top             =   3060
               Width           =   2175
            End
            Begin VB.Label Label42 
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
               Left            =   3480
               TabIndex        =   131
               Top             =   2580
               Width           =   2055
            End
         End
         Begin TabDlg.SSTab SSTab5 
            Height          =   6315
            Left            =   1920
            TabIndex        =   97
            Top             =   1080
            Width           =   14415
            _ExtentX        =   25426
            _ExtentY        =   11139
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
            TabCaption(0)   =   "Shto Ilaçe"
            TabPicture(0)   =   "Admin.frx":50A06
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "Label33(0)"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).Control(1)=   "Label34(0)"
            Tab(0).Control(1).Enabled=   0   'False
            Tab(0).Control(2)=   "Label35(0)"
            Tab(0).Control(2).Enabled=   0   'False
            Tab(0).Control(3)=   "Command18"
            Tab(0).Control(3).Enabled=   0   'False
            Tab(0).Control(4)=   "Command19"
            Tab(0).Control(4).Enabled=   0   'False
            Tab(0).Control(5)=   "Command20"
            Tab(0).Control(5).Enabled=   0   'False
            Tab(0).Control(6)=   "Command21"
            Tab(0).Control(6).Enabled=   0   'False
            Tab(0).Control(7)=   "Text27"
            Tab(0).Control(7).Enabled=   0   'False
            Tab(0).Control(8)=   "Text28"
            Tab(0).Control(8).Enabled=   0   'False
            Tab(0).Control(9)=   "Text29"
            Tab(0).Control(9).Enabled=   0   'False
            Tab(0).ControlCount=   10
            TabCaption(1)   =   "Ndrysho Ilaçe"
            TabPicture(1)   =   "Admin.frx":50A22
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "Command22"
            Tab(1).Control(1)=   "Command23"
            Tab(1).Control(2)=   "Command24"
            Tab(1).Control(3)=   "Command25"
            Tab(1).Control(4)=   "Command26"
            Tab(1).Control(5)=   "Frame5"
            Tab(1).Control(6)=   "Adodc3"
            Tab(1).Control(7)=   "Frame6"
            Tab(1).ControlCount=   8
            Begin VB.Frame Frame6 
               Caption         =   "Ndrysho Ilaçe"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   4215
               Left            =   -66840
               TabIndex        =   117
               Top             =   1380
               Width           =   5775
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
                  ItemData        =   "Admin.frx":50A3E
                  Left            =   2160
                  List            =   "Admin.frx":50A48
                  TabIndex        =   125
                  Top             =   2760
                  Width           =   2895
               End
               Begin VB.TextBox Text32 
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
                  Left            =   2160
                  TabIndex        =   124
                  Top             =   2280
                  Width           =   2895
               End
               Begin VB.TextBox Text31 
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
                  Left            =   2160
                  TabIndex        =   123
                  Top             =   1800
                  Width           =   2895
               End
               Begin VB.TextBox Text30 
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
                  Left            =   2160
                  TabIndex        =   122
                  Top             =   1320
                  Width           =   2895
               End
               Begin VB.Label Label40 
                  Caption         =   "Gjendja"
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
                  Left            =   900
                  TabIndex        =   121
                  Top             =   2760
                  Width           =   975
               End
               Begin VB.Label Label37 
                  Caption         =   "Kodi i Ilaçit"
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
                  TabIndex        =   120
                  Top             =   1320
                  Width           =   1695
               End
               Begin VB.Label Label38 
                  Caption         =   "Emri i Ilacit"
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
                  TabIndex        =   119
                  Top             =   1800
                  Width           =   1575
               End
               Begin VB.Label Label39 
                  Caption         =   "Çmimi i Ilacit"
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
                  Left            =   120
                  TabIndex        =   118
                  Top             =   2280
                  Width           =   1815
               End
            End
            Begin MSAdodcLib.Adodc Adodc3 
               Height          =   495
               Left            =   -74640
               Top             =   420
               Visible         =   0   'False
               Width           =   1935
               _ExtentX        =   3413
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
               Connect         =   $"Admin.frx":50A66
               OLEDBString     =   $"Admin.frx":50B05
               OLEDBFile       =   ""
               DataSourceName  =   ""
               OtherAttributes =   ""
               UserName        =   ""
               Password        =   ""
               RecordSource    =   "Select * From Ilace"
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
            Begin VB.Frame Frame5 
               Caption         =   "Kerko Ilaçe"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   4215
               Left            =   -74640
               TabIndex        =   113
               Top             =   1380
               Width           =   7575
               Begin MSDataGridLib.DataGrid DataGrid3 
                  Bindings        =   "Admin.frx":50BA4
                  Height          =   1935
                  Left            =   240
                  TabIndex        =   116
                  Top             =   1920
                  Width           =   7095
                  _ExtentX        =   12515
                  _ExtentY        =   3413
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
                     DataField       =   "IdIlace"
                     Caption         =   "IdIlace"
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
                     DataField       =   "EmerIlace"
                     Caption         =   "EmerIlace"
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
                     DataField       =   "CmimiIlace"
                     Caption         =   "CmimiIlace"
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
                     DataField       =   "Gjendje"
                     Caption         =   "Gjendje"
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
                  Left            =   2160
                  TabIndex        =   115
                  Top             =   840
                  Width           =   3255
               End
               Begin VB.Label Label36 
                  Caption         =   "Emer Ilaçi"
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
                  Left            =   720
                  TabIndex        =   114
                  Top             =   840
                  Width           =   1335
               End
            End
            Begin VB.CommandButton Command26 
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
               Left            =   -66360
               Picture         =   "Admin.frx":50BB9
               Style           =   1  'Graphical
               TabIndex        =   112
               Top             =   420
               Width           =   855
            End
            Begin VB.CommandButton Command25 
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
               Left            =   -67320
               Picture         =   "Admin.frx":5429B
               Style           =   1  'Graphical
               TabIndex        =   111
               Top             =   420
               Width           =   855
            End
            Begin VB.CommandButton Command24 
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
               Left            =   -68280
               Picture         =   "Admin.frx":576DC
               Style           =   1  'Graphical
               TabIndex        =   110
               Top             =   420
               Width           =   855
            End
            Begin VB.CommandButton Command23 
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
               Left            =   -69240
               Picture         =   "Admin.frx":5A9AF
               Style           =   1  'Graphical
               TabIndex        =   109
               Top             =   420
               Width           =   855
            End
            Begin VB.CommandButton Command22 
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
               Left            =   -70200
               Picture         =   "Admin.frx":5DB76
               Style           =   1  'Graphical
               TabIndex        =   108
               Top             =   420
               Width           =   855
            End
            Begin VB.TextBox Text29 
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
               Left            =   5760
               TabIndex        =   107
               Top             =   3180
               Width           =   2775
            End
            Begin VB.TextBox Text28 
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
               Left            =   5760
               TabIndex        =   106
               Top             =   2700
               Width           =   2775
            End
            Begin VB.TextBox Text27 
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
               Left            =   5760
               TabIndex        =   105
               Top             =   2220
               Width           =   2775
            End
            Begin VB.CommandButton Command21 
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
               Left            =   8400
               Picture         =   "Admin.frx":60E61
               Style           =   1  'Graphical
               TabIndex        =   101
               Top             =   420
               Width           =   855
            End
            Begin VB.CommandButton Command20 
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
               Left            =   7440
               Picture         =   "Admin.frx":64543
               Style           =   1  'Graphical
               TabIndex        =   100
               Top             =   420
               Width           =   855
            End
            Begin VB.CommandButton Command19 
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
               Left            =   6480
               Picture         =   "Admin.frx":6770A
               Style           =   1  'Graphical
               TabIndex        =   99
               Top             =   420
               Width           =   855
            End
            Begin VB.CommandButton Command18 
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
               Left            =   5520
               Picture         =   "Admin.frx":6A782
               Style           =   1  'Graphical
               TabIndex        =   98
               Top             =   420
               Width           =   855
            End
            Begin VB.Label Label35 
               Caption         =   "Çmimi i Ilacit"
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
               Index           =   0
               Left            =   3600
               TabIndex        =   104
               Top             =   3180
               Width           =   1815
            End
            Begin VB.Label Label34 
               Caption         =   "Emri i Ilacit"
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
               Index           =   0
               Left            =   3840
               TabIndex        =   103
               Top             =   2700
               Width           =   1575
            End
            Begin VB.Label Label33 
               Caption         =   "Kodi i Ilaçit"
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
               Index           =   0
               Left            =   3840
               TabIndex        =   102
               Top             =   2220
               Width           =   1695
            End
         End
      End
      Begin TabDlg.SSTab SSTab3 
         Height          =   8295
         Left            =   -72240
         TabIndex        =   74
         Top             =   1440
         Width           =   14415
         _ExtentX        =   25426
         _ExtentY        =   14631
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
         TabCaption(0)   =   "Shto Departament"
         TabPicture(0)   =   "Admin.frx":6DA6D
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Label28"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Label29"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "Text23"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "Text24"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "Command10"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "Command11"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "Command12"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).Control(7)=   "Command13"
         Tab(0).Control(7).Enabled=   0   'False
         Tab(0).ControlCount=   8
         TabCaption(1)   =   "Ndrysho Departament"
         TabPicture(1)   =   "Admin.frx":6DA89
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Command17"
         Tab(1).Control(1)=   "Adodc2"
         Tab(1).Control(2)=   "Command16"
         Tab(1).Control(3)=   "Command15"
         Tab(1).Control(4)=   "Command14"
         Tab(1).Control(5)=   "Frame4"
         Tab(1).Control(6)=   "Frame3"
         Tab(1).Control(7)=   "DataGrid2"
         Tab(1).ControlCount=   8
         Begin VB.CommandButton Command17 
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
            Left            =   -67080
            Picture         =   "Admin.frx":6DAA5
            Style           =   1  'Graphical
            TabIndex        =   95
            Top             =   720
            Width           =   855
         End
         Begin MSAdodcLib.Adodc Adodc2 
            Height          =   495
            Left            =   -74520
            Top             =   960
            Visible         =   0   'False
            Width           =   2295
            _ExtentX        =   4048
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
            Connect         =   $"Admin.frx":71187
            OLEDBString     =   $"Admin.frx":71226
            OLEDBFile       =   ""
            DataSourceName  =   ""
            OtherAttributes =   ""
            UserName        =   ""
            Password        =   ""
            RecordSource    =   "Select * From Departamenti"
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
         Begin VB.CommandButton Command16 
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
            Left            =   -68040
            Picture         =   "Admin.frx":712C5
            Style           =   1  'Graphical
            TabIndex        =   88
            Top             =   720
            Width           =   855
         End
         Begin VB.CommandButton Command15 
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
            Left            =   -69000
            Picture         =   "Admin.frx":74706
            Style           =   1  'Graphical
            TabIndex        =   87
            Top             =   720
            Width           =   855
         End
         Begin VB.CommandButton Command14 
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
            Left            =   -69960
            Picture         =   "Admin.frx":778CD
            Style           =   1  'Graphical
            TabIndex        =   86
            Top             =   720
            Width           =   855
         End
         Begin VB.Frame Frame4 
            Caption         =   "Ndrysho"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1695
            Left            =   -67920
            TabIndex        =   85
            Top             =   2280
            Width           =   6495
            Begin VB.TextBox Text26 
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
               Left            =   3120
               TabIndex        =   94
               Top             =   960
               Width           =   2535
            End
            Begin VB.TextBox Text25 
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
               Height          =   405
               Left            =   3120
               TabIndex        =   93
               Top             =   480
               Width           =   2535
            End
            Begin VB.Label Label32 
               Caption         =   "Id Departamentit"
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
               Left            =   800
               TabIndex        =   92
               Top             =   480
               Width           =   2175
            End
            Begin VB.Label Label31 
               Caption         =   "Emri i Departamentit"
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
               TabIndex        =   91
               Top             =   960
               Width           =   2655
            End
         End
         Begin VB.Frame Frame3 
            Caption         =   "Kerko"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1695
            Left            =   -74160
            TabIndex        =   84
            Top             =   2280
            Width           =   5895
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
               Left            =   3000
               TabIndex        =   90
               Top             =   720
               Width           =   2775
            End
            Begin VB.Label Label30 
               Caption         =   "Emri i Departamentit"
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
               Left            =   120
               TabIndex        =   89
               Top             =   720
               Width           =   2655
            End
         End
         Begin MSDataGridLib.DataGrid DataGrid2 
            Bindings        =   "Admin.frx":7ABB8
            Height          =   3255
            Left            =   -72960
            TabIndex        =   83
            Top             =   4320
            Width           =   10215
            _ExtentX        =   18018
            _ExtentY        =   5741
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
            ColumnCount     =   2
            BeginProperty Column00 
               DataField       =   "IdDepartamenti"
               Caption         =   "IdDepartamenti"
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
               DataField       =   "EmriDepartamentit"
               Caption         =   "EmriDepartamentit"
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
            EndProperty
         End
         Begin VB.CommandButton Command13 
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
            Left            =   7920
            Picture         =   "Admin.frx":7ABCD
            Style           =   1  'Graphical
            TabIndex        =   82
            Top             =   720
            Width           =   855
         End
         Begin VB.CommandButton Command12 
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
            Left            =   6960
            Picture         =   "Admin.frx":7E2AF
            Style           =   1  'Graphical
            TabIndex        =   81
            Top             =   720
            Width           =   855
         End
         Begin VB.CommandButton Command11 
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
            Left            =   6000
            Picture         =   "Admin.frx":81476
            Style           =   1  'Graphical
            TabIndex        =   80
            Top             =   720
            Width           =   855
         End
         Begin VB.CommandButton Command10 
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
            Left            =   5040
            Picture         =   "Admin.frx":844EE
            Style           =   1  'Graphical
            TabIndex        =   79
            Top             =   720
            Width           =   855
         End
         Begin VB.TextBox Text24 
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
            Left            =   5520
            TabIndex        =   78
            Top             =   3720
            Width           =   3495
         End
         Begin VB.TextBox Text23 
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
            Left            =   5520
            TabIndex        =   77
            Top             =   3000
            Width           =   3495
         End
         Begin VB.Label Label29 
            Caption         =   "Emer Departamenti"
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
            Left            =   2850
            TabIndex        =   76
            Top             =   3720
            Width           =   2535
         End
         Begin VB.Label Label28 
            Caption         =   "Id Departamenti"
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
            Left            =   3240
            TabIndex        =   75
            Top             =   3000
            Width           =   2055
         End
      End
      Begin TabDlg.SSTab SSTab2 
         Height          =   8295
         Left            =   480
         TabIndex        =   1
         Top             =   1440
         Width           =   19215
         _ExtentX        =   33893
         _ExtentY        =   14631
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
         TabCaption(0)   =   "Shto Mjek"
         TabPicture(0)   =   "Admin.frx":877D9
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Label1"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Label2"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "Label3"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "Label4"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "Label5"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "Label6"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "Label7"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).Control(7)=   "Label8"
         Tab(0).Control(7).Enabled=   0   'False
         Tab(0).Control(8)=   "Label9"
         Tab(0).Control(8).Enabled=   0   'False
         Tab(0).Control(9)=   "Label10"
         Tab(0).Control(9).Enabled=   0   'False
         Tab(0).Control(10)=   "Label11"
         Tab(0).Control(10).Enabled=   0   'False
         Tab(0).Control(11)=   "Label12"
         Tab(0).Control(11).Enabled=   0   'False
         Tab(0).Control(12)=   "Label41"
         Tab(0).Control(12).Enabled=   0   'False
         Tab(0).Control(13)=   "Label58"
         Tab(0).Control(13).Enabled=   0   'False
         Tab(0).Control(14)=   "Text1"
         Tab(0).Control(14).Enabled=   0   'False
         Tab(0).Control(15)=   "Text2"
         Tab(0).Control(15).Enabled=   0   'False
         Tab(0).Control(16)=   "Text3"
         Tab(0).Control(16).Enabled=   0   'False
         Tab(0).Control(17)=   "Text4"
         Tab(0).Control(17).Enabled=   0   'False
         Tab(0).Control(18)=   "Text5"
         Tab(0).Control(18).Enabled=   0   'False
         Tab(0).Control(19)=   "Text6"
         Tab(0).Control(19).Enabled=   0   'False
         Tab(0).Control(20)=   "Text7"
         Tab(0).Control(20).Enabled=   0   'False
         Tab(0).Control(21)=   "Combo1"
         Tab(0).Control(21).Enabled=   0   'False
         Tab(0).Control(22)=   "Combo2"
         Tab(0).Control(22).Enabled=   0   'False
         Tab(0).Control(23)=   "Combo3"
         Tab(0).Control(23).Enabled=   0   'False
         Tab(0).Control(24)=   "Option1"
         Tab(0).Control(24).Enabled=   0   'False
         Tab(0).Control(25)=   "Option2"
         Tab(0).Control(25).Enabled=   0   'False
         Tab(0).Control(26)=   "DTPicker1"
         Tab(0).Control(26).Enabled=   0   'False
         Tab(0).Control(27)=   "DTPicker2"
         Tab(0).Control(27).Enabled=   0   'False
         Tab(0).Control(28)=   "Text8"
         Tab(0).Control(28).Enabled=   0   'False
         Tab(0).Control(29)=   "Text9"
         Tab(0).Control(29).Enabled=   0   'False
         Tab(0).Control(30)=   "Text10"
         Tab(0).Control(30).Enabled=   0   'False
         Tab(0).Control(31)=   "Command1"
         Tab(0).Control(31).Enabled=   0   'False
         Tab(0).Control(32)=   "Command2"
         Tab(0).Control(32).Enabled=   0   'False
         Tab(0).Control(33)=   "Command3"
         Tab(0).Control(33).Enabled=   0   'False
         Tab(0).Control(34)=   "Command4"
         Tab(0).Control(34).Enabled=   0   'False
         Tab(0).Control(35)=   "Text15"
         Tab(0).Control(35).Enabled=   0   'False
         Tab(0).Control(36)=   "Text16"
         Tab(0).Control(36).Enabled=   0   'False
         Tab(0).Control(37)=   "Command49"
         Tab(0).Control(37).Enabled=   0   'False
         Tab(0).ControlCount=   38
         TabCaption(1)   =   "Ndrysho Mjek"
         TabPicture(1)   =   "Admin.frx":877F5
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Command5"
         Tab(1).Control(1)=   "Command6"
         Tab(1).Control(2)=   "Command7"
         Tab(1).Control(3)=   "Command8"
         Tab(1).Control(4)=   "Command9"
         Tab(1).Control(5)=   "Picture1"
         Tab(1).Control(6)=   "Adodc1"
         Tab(1).ControlCount=   7
         Begin VB.CommandButton Command49 
            Caption         =   "Gjenero Id Mjeku"
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
            Left            =   3840
            TabIndex        =   201
            Top             =   1320
            Width           =   2415
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
            Height          =   435
            IMEMode         =   3  'DISABLE
            Left            =   13440
            PasswordChar    =   "*"
            TabIndex        =   195
            Top             =   3240
            Width           =   1935
         End
         Begin VB.TextBox Text15 
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
            Left            =   8400
            TabIndex        =   194
            Top             =   3300
            Width           =   1935
         End
         Begin MSAdodcLib.Adodc Adodc1 
            Height          =   330
            Left            =   -74520
            Top             =   660
            Visible         =   0   'False
            Width           =   3495
            _ExtentX        =   6165
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
            Connect         =   $"Admin.frx":87811
            OLEDBString     =   $"Admin.frx":878B0
            OLEDBFile       =   ""
            DataSourceName  =   ""
            OtherAttributes =   ""
            UserName        =   ""
            Password        =   ""
            RecordSource    =   $"Admin.frx":8794F
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
         Begin VB.PictureBox Picture1 
            Height          =   6495
            Left            =   -74880
            ScaleHeight     =   6435
            ScaleWidth      =   18915
            TabIndex        =   40
            Top             =   1260
            Width           =   18975
            Begin MSDataGridLib.DataGrid DataGrid1 
               Bindings        =   "Admin.frx":87B00
               Height          =   3375
               Left            =   120
               TabIndex        =   47
               Top             =   3000
               Width           =   18855
               _ExtentX        =   33258
               _ExtentY        =   5953
               _Version        =   393216
               AllowUpdate     =   0   'False
               Enabled         =   -1  'True
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
               ColumnCount     =   15
               BeginProperty Column00 
                  DataField       =   "IdMjek"
                  Caption         =   "IdMjek"
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
                  DataField       =   "llojKualifikim"
                  Caption         =   "llojKualifikim"
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
                  DataField       =   "Email"
                  Caption         =   "Email"
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
                  DataField       =   "llojSpecializimi"
                  Caption         =   "llojSpecializimi"
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
                  DataField       =   "EmriDepartamentit"
                  Caption         =   "EmriDepartamentit"
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
                  DataField       =   "Dt_Punesimit"
                  Caption         =   "Dt_Punesimit"
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
                  DataField       =   "StatusPune"
                  Caption         =   "StatusPune"
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
                  DataField       =   "Perdoruesi"
                  Caption         =   "Perdoruesi"
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
               BeginProperty Column14 
                  DataField       =   "Fjalekalimi"
                  Caption         =   "Fjalekalimi"
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
                     ColumnWidth     =   1140.095
                  EndProperty
                  BeginProperty Column12 
                     ColumnWidth     =   1739.906
                  EndProperty
                  BeginProperty Column13 
                     ColumnWidth     =   1739.906
                  EndProperty
                  BeginProperty Column14 
                     ColumnWidth     =   1739.906
                  EndProperty
               EndProperty
            End
            Begin VB.Frame Frame2 
               Caption         =   "Frame2"
               Height          =   2655
               Left            =   6000
               TabIndex        =   46
               Top             =   0
               Width           =   12855
               Begin VB.TextBox Text33 
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
                  Left            =   10440
                  PasswordChar    =   "*"
                  TabIndex        =   200
                  Top             =   2160
                  Width           =   1935
               End
               Begin VB.TextBox Text21 
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
                  Left            =   10440
                  TabIndex        =   199
                  Top             =   1680
                  Width           =   1935
               End
               Begin VB.ComboBox Combo18 
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
                  ItemData        =   "Admin.frx":87B15
                  Left            =   1560
                  List            =   "Admin.frx":87B1F
                  TabIndex        =   192
                  Top             =   2160
                  Width           =   2175
               End
               Begin MSComCtl2.DTPicker DTPicker4 
                  Height          =   375
                  Left            =   10440
                  TabIndex        =   186
                  Top             =   720
                  Width           =   1935
                  _ExtentX        =   3413
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
                  Format          =   96600067
                  CurrentDate     =   42594
               End
               Begin MSComCtl2.DTPicker DTPicker3 
                  Height          =   375
                  Left            =   5520
                  TabIndex        =   185
                  Top             =   240
                  Width           =   2055
                  _ExtentX        =   3625
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
                  Format          =   96600067
                  CurrentDate     =   42594
               End
               Begin VB.TextBox Text22 
                  Height          =   285
                  Left            =   12600
                  TabIndex        =   73
                  Text            =   "Text22"
                  Top             =   240
                  Visible         =   0   'False
                  Width           =   255
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
                  ItemData        =   "Admin.frx":87B29
                  Left            =   10440
                  List            =   "Admin.frx":87B33
                  TabIndex        =   72
                  Top             =   1200
                  Width           =   1935
               End
               Begin VB.TextBox Text20 
                  Height          =   405
                  Left            =   7800
                  TabIndex        =   71
                  Text            =   "Text20"
                  Top             =   2160
                  Visible         =   0   'False
                  Width           =   495
               End
               Begin VB.ComboBox Combo8 
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
                  Left            =   10440
                  TabIndex        =   70
                  Top             =   240
                  Width           =   1935
               End
               Begin VB.ComboBox Combo7 
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
                  Left            =   5520
                  TabIndex        =   69
                  Top             =   2160
                  Width           =   2175
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
                  Left            =   5520
                  TabIndex        =   68
                  Top             =   720
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
                  Left            =   5520
                  TabIndex        =   67
                  Top             =   1680
                  Width           =   2175
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
                  Height          =   375
                  Left            =   5520
                  TabIndex        =   66
                  Top             =   1200
                  Width           =   2175
               End
               Begin VB.TextBox Text17 
                  Height          =   375
                  Left            =   7800
                  TabIndex        =   65
                  Text            =   "Text17"
                  Top             =   720
                  Visible         =   0   'False
                  Width           =   375
               End
               Begin VB.TextBox Text14 
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
                  Left            =   1560
                  TabIndex        =   63
                  Top             =   1680
                  Width           =   2175
               End
               Begin VB.TextBox Text13 
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
                  Left            =   1560
                  TabIndex        =   62
                  Top             =   1200
                  Width           =   2175
               End
               Begin VB.TextBox Text12 
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
                  Left            =   1560
                  TabIndex        =   61
                  Top             =   720
                  Width           =   2175
               End
               Begin VB.TextBox Text11 
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
                  Left            =   1560
                  TabIndex        =   60
                  Top             =   240
                  Width           =   2175
               End
               Begin VB.Label Label60 
                  Caption         =   "Id Perdoruesi"
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
                  Left            =   8640
                  TabIndex        =   198
                  Top             =   1680
                  Width           =   1695
               End
               Begin VB.Label Label59 
                  Caption         =   "Fjalekalimi"
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
                  Left            =   8760
                  TabIndex        =   197
                  Top             =   2160
                  Width           =   1575
               End
               Begin VB.Label Label27 
                  Caption         =   "Statusi i Punes"
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
                  Left            =   8445
                  TabIndex        =   64
                  Top             =   1200
                  Width           =   1935
               End
               Begin VB.Label Label26 
                  Caption         =   "Id Mjekut"
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
                  TabIndex        =   59
                  Top             =   240
                  Width           =   1215
               End
               Begin VB.Label Label25 
                  Caption         =   "Data Punesimit"
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
                  Left            =   8385
                  TabIndex        =   58
                  Top             =   720
                  Width           =   1935
               End
               Begin VB.Label Label24 
                  Caption         =   "Emri Departamentit"
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
                  Left            =   7800
                  TabIndex        =   57
                  Top             =   240
                  Width           =   2535
               End
               Begin VB.Label Label23 
                  Caption         =   "Specializimi"
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
                  Left            =   3840
                  TabIndex        =   56
                  Top             =   2160
                  Width           =   1575
               End
               Begin VB.Label Label22 
                  Caption         =   "Email"
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
                  TabIndex        =   55
                  Top             =   1680
                  Width           =   735
               End
               Begin VB.Label Label21 
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
                  Left            =   4365
                  TabIndex        =   54
                  Top             =   1200
                  Width           =   1095
               End
               Begin VB.Label Label20 
                  Caption         =   "Kualifikimi"
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
                  Left            =   3915
                  TabIndex        =   53
                  Top             =   720
                  Width           =   1455
               End
               Begin VB.Label Label19 
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
                  Left            =   4080
                  TabIndex        =   52
                  Top             =   240
                  Width           =   1335
               End
               Begin VB.Label Label18 
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
                  Left            =   555
                  TabIndex        =   51
                  Top             =   2160
                  Width           =   855
               End
               Begin VB.Label Label17 
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
                  Left            =   240
                  TabIndex        =   50
                  Top             =   1680
                  Width           =   1095
               End
               Begin VB.Label Label16 
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
                  Left            =   600
                  TabIndex        =   49
                  Top             =   1200
                  Width           =   855
               End
               Begin VB.Label Label15 
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
                  Left            =   700
                  TabIndex        =   48
                  Top             =   720
                  Width           =   615
               End
            End
            Begin VB.Frame Frame1 
               Caption         =   "Kerko"
               Height          =   1695
               Left            =   120
               TabIndex        =   41
               Top             =   0
               Width           =   5775
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
                  Left            =   2640
                  TabIndex        =   43
                  Top             =   840
                  Width           =   2535
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
                  Left            =   2640
                  TabIndex        =   42
                  Top             =   360
                  Width           =   2535
               End
               Begin VB.Label Label14 
                  Caption         =   "Mbiemri i Mjekut"
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
                  TabIndex        =   45
                  Top             =   840
                  Width           =   2415
               End
               Begin VB.Label Label13 
                  Caption         =   "Emri i Mjekut"
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
                  TabIndex        =   44
                  Top             =   360
                  Width           =   1815
               End
            End
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
            Left            =   -64800
            Picture         =   "Admin.frx":87B45
            Style           =   1  'Graphical
            TabIndex        =   39
            Top             =   420
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
            Left            =   -65760
            Picture         =   "Admin.frx":8B227
            Style           =   1  'Graphical
            TabIndex        =   38
            Top             =   420
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
            Left            =   -66720
            Picture         =   "Admin.frx":8E668
            Style           =   1  'Graphical
            TabIndex        =   37
            Top             =   420
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
            Left            =   -67680
            Picture         =   "Admin.frx":9193B
            Style           =   1  'Graphical
            TabIndex        =   36
            Top             =   420
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
            Left            =   -68640
            Picture         =   "Admin.frx":94B02
            Style           =   1  'Graphical
            TabIndex        =   35
            Top             =   420
            Width           =   855
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
            Left            =   10200
            Picture         =   "Admin.frx":97DED
            Style           =   1  'Graphical
            TabIndex        =   34
            Top             =   420
            Width           =   735
         End
         Begin VB.CommandButton Command3 
            Caption         =   "Riruaj"
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
            Left            =   9240
            Picture         =   "Admin.frx":9B4CF
            Style           =   1  'Graphical
            TabIndex        =   33
            Top             =   420
            Width           =   735
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
            Left            =   8280
            Picture         =   "Admin.frx":9E696
            Style           =   1  'Graphical
            TabIndex        =   32
            Top             =   420
            Width           =   735
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
            Left            =   7320
            Picture         =   "Admin.frx":A170E
            Style           =   1  'Graphical
            TabIndex        =   31
            Top             =   420
            Width           =   735
         End
         Begin VB.TextBox Text10 
            Height          =   285
            Left            =   15960
            TabIndex        =   30
            Top             =   2340
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.TextBox Text9 
            Height          =   285
            Left            =   15960
            TabIndex        =   29
            Top             =   1860
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.TextBox Text8 
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
            Left            =   8400
            TabIndex        =   28
            Top             =   2820
            Width           =   1935
         End
         Begin MSComCtl2.DTPicker DTPicker2 
            Bindings        =   "Admin.frx":A49F9
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "yyyy/MM/dd"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
            DataSource      =   "(None)"
            Height          =   375
            Left            =   13440
            TabIndex        =   27
            Top             =   2820
            Width           =   1935
            _ExtentX        =   3413
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
            Format          =   96600067
            CurrentDate     =   42593
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "yyyy/MM/dd"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   3
            EndProperty
            Height          =   375
            Left            =   3840
            TabIndex        =   26
            Top             =   4260
            Width           =   2295
            _ExtentX        =   4048
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
            Format          =   96600067
            CurrentDate     =   42593
            MinDate         =   -109184
         End
         Begin VB.OptionButton Option2 
            Caption         =   "F"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   4560
            TabIndex        =   25
            Top             =   3780
            Width           =   735
         End
         Begin VB.OptionButton Option1 
            Caption         =   "M"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   3840
            TabIndex        =   24
            Top             =   3780
            Width           =   735
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
            Left            =   13440
            TabIndex        =   11
            Top             =   2340
            Width           =   1935
         End
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
            Left            =   13440
            TabIndex        =   10
            Top             =   1860
            Width           =   1935
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
            Left            =   8400
            TabIndex        =   9
            Top             =   1860
            Width           =   1935
         End
         Begin VB.TextBox Text7 
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
            Left            =   8400
            TabIndex        =   8
            Top             =   2340
            Width           =   1935
         End
         Begin VB.TextBox Text6 
            Height          =   285
            Left            =   10440
            TabIndex        =   7
            Top             =   1740
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.TextBox Text5 
            Height          =   375
            Left            =   6000
            TabIndex        =   6
            Top             =   3780
            Visible         =   0   'False
            Width           =   375
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
            Height          =   435
            Left            =   3840
            TabIndex        =   5
            Top             =   3300
            Width           =   2415
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
            Height          =   435
            Left            =   3840
            TabIndex        =   4
            Top             =   2820
            Width           =   2415
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
            Height          =   435
            Left            =   3840
            TabIndex        =   3
            Top             =   2340
            Width           =   2415
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
            Height          =   435
            Left            =   3840
            TabIndex        =   2
            Top             =   1860
            Width           =   2415
         End
         Begin VB.Label Label58 
            Caption         =   "Fjalekalimi"
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
            Left            =   11880
            TabIndex        =   196
            Top             =   3300
            Width           =   1575
         End
         Begin VB.Label Label41 
            Caption         =   "Id Perdoruesi"
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
            Left            =   6480
            TabIndex        =   193
            Top             =   3300
            Width           =   1695
         End
         Begin VB.Label Label12 
            Caption         =   "Data Punesimit"
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
            Left            =   11400
            TabIndex        =   23
            Top             =   2820
            Width           =   2055
         End
         Begin VB.Label Label11 
            Caption         =   "Emri Departamentit"
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
            Left            =   10875
            TabIndex        =   22
            Top             =   2340
            Width           =   2535
         End
         Begin VB.Label Label10 
            Caption         =   "Specializimi"
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
            Left            =   11760
            TabIndex        =   21
            Top             =   1860
            Width           =   1575
         End
         Begin VB.Label Label9 
            Caption         =   "E-mail"
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
            Left            =   7260
            TabIndex        =   20
            Top             =   2820
            Width           =   975
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
            Left            =   7080
            TabIndex        =   19
            Top             =   2340
            Width           =   1095
         End
         Begin VB.Label Label7 
            Caption         =   "Kualifikimi"
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
            Left            =   6600
            TabIndex        =   18
            Top             =   1860
            Width           =   1575
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
            Height          =   375
            Left            =   2040
            TabIndex        =   17
            Top             =   4260
            Width           =   1335
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
            Left            =   2520
            TabIndex        =   16
            Top             =   3780
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
            Left            =   2280
            TabIndex        =   15
            Top             =   3300
            Width           =   1095
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
            Left            =   2520
            TabIndex        =   14
            Top             =   2820
            Width           =   855
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
            Left            =   2640
            TabIndex        =   13
            Top             =   2340
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "Id Mjeku"
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
            Left            =   2160
            TabIndex        =   12
            Top             =   1860
            Width           =   1215
         End
      End
   End
   Begin VB.Menu mnuMjek 
      Caption         =   "&Mjeku"
      WindowList      =   -1  'True
      Begin VB.Menu shtoMjek 
         Caption         =   "Shto Mjek"
      End
      Begin VB.Menu ndryshoMjek 
         Caption         =   "Ndrysho Mjek"
      End
      Begin VB.Menu raportMjek 
         Caption         =   "Raport"
      End
   End
   Begin VB.Menu mnuDepartament 
      Caption         =   "&Departament"
      Begin VB.Menu shtoDept 
         Caption         =   "Shto Departament"
      End
      Begin VB.Menu ndryshoDept 
         Caption         =   "Ndrysho Departament"
      End
   End
   Begin VB.Menu mnuIlace 
      Caption         =   "&Ilace"
      Begin VB.Menu shtoIlace 
         Caption         =   "Shto Ilace"
      End
      Begin VB.Menu ndryshoIlace 
         Caption         =   "Ndrysho Ilace"
      End
      Begin VB.Menu raportIlace 
         Caption         =   "Raport"
      End
   End
   Begin VB.Menu mnuAnaliza 
      Caption         =   "&Analiza"
      Begin VB.Menu shtoAnaliza 
         Caption         =   "Shto Analiza"
      End
      Begin VB.Menu ndryshoAnaliza 
         Caption         =   "Ndrysho Analiza"
      End
      Begin VB.Menu raportAnaliza 
         Caption         =   "Raport"
      End
   End
   Begin VB.Menu mnuInjeksione 
      Caption         =   "&Injeksione"
      Begin VB.Menu shtoInjeksione 
         Caption         =   "Shto Injeksione"
      End
      Begin VB.Menu ndryshoInjeksione 
         Caption         =   "Ndrysho Injeksione"
      End
      Begin VB.Menu raportInjeksione 
         Caption         =   "Raport"
      End
   End
End
Attribute VB_Name = "Admin"
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
Dim status As String
Dim status1 As String
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
'======================================================= MENUTE =======================================================================


'============MENU MJEK====

Private Sub shtoMjek_Click()
Admin.SSTab1.Tab = 0
Admin.SSTab2.Tab = 0
End Sub

Private Sub ndryshoMjek_Click()
Admin.SSTab1.Tab = 0
Admin.SSTab2.Tab = 1
End Sub

Private Sub raportMjek_Click()
Admin.SSTab1.Tab = 0
Admin.SSTab2.Tab = 2
MjekuRaport.Show
End Sub

'============MENU DEPARTAMENT====

Private Sub shtoDept_Click()
Admin.SSTab1.Tab = 1
Admin.SSTab3.Tab = 0
End Sub

Private Sub ndryshoDept_Click()
Admin.SSTab1.Tab = 1
Admin.SSTab3.Tab = 1
End Sub

'============MENU ILACE====

Private Sub shtoIlace_Click()
Admin.SSTab1.Tab = 2
Admin.SSTab4.Tab = 0
Admin.SSTab5.Tab = 0
End Sub

Private Sub ndryshoIlace_Click()
Admin.SSTab1.Tab = 2
Admin.SSTab4.Tab = 0
Admin.SSTab5.Tab = 1
End Sub

Private Sub raportIlace_Click()
Admin.SSTab1.Tab = 2
Admin.SSTab4.Tab = 0
Admin.SSTab5.Tab = 2
IlaceRaport.Show
End Sub

'============MENU ANALIZA====

Private Sub shtoAnaliza_Click()
Admin.SSTab1.Tab = 2
Admin.SSTab4.Tab = 1
Admin.SSTab6.Tab = 0
End Sub

Private Sub ndryshoAnaliza_Click()
Admin.SSTab1.Tab = 2
Admin.SSTab4.Tab = 1
Admin.SSTab6.Tab = 1
End Sub

Private Sub raportAnaliza_Click()
Admin.SSTab1.Tab = 2
Admin.SSTab4.Tab = 1
Admin.SSTab6.Tab = 2
AnalizaRaport.Show
End Sub

'============MENU INJEKSIONE====

Private Sub shtoInjeksione_Click()
Admin.SSTab1.Tab = 2
Admin.SSTab4.Tab = 2
Admin.SSTab7.Tab = 0
End Sub

Private Sub ndryshoInjeksione_Click()
Admin.SSTab1.Tab = 2
Admin.SSTab4.Tab = 2
Admin.SSTab7.Tab = 1
End Sub

Private Sub raportInjeksione_Click()
Admin.SSTab1.Tab = 2
Admin.SSTab4.Tab = 2
Admin.SSTab7.Tab = 2
InjeksioneRaport.Show
End Sub

'=======================================================SHTO MJEK====================================================================================


'Selektimi i te dhenave nga tabela Kualifikimi dhe vendosja e te dhenave tek Combobox
Private Sub Combo1_Add()
Combo1.Clear
Dim strng As String
Con = "Provider=MSDASQL.1;Data Source=Juli;Initial Catalog=SQLEXPRESS"
Con.Open "Driver=SQL Server;Server=.\SQLEXPRESS;Database=Hospital;uid=;pwd=;"
rec.Open "select llojKualifikim from Kualifikimi", Con, adOpenUnspecified, adLockReadOnly
Combo1.AddItem "Selekto"
Combo1.Text = Me.Combo1.List(0)
While Not rec.EOF
Combo1.AddItem rec(0)
rec.MoveNext
Wend
Con.Close
End Sub

Private Sub Combo1_Click()
Dim strng As String
Con = "Provider=MSDASQL.1;Data Source=Juli;Initial Catalog=SQLEXPRESS"
Con.Open "Driver=SQL Server;Server=.\SQLEXPRESS;Database=Hospital;uid=;pwd=;"
rec.Open "select idKualifikim from Kualifikimi Where llojKualifikim = '" & Combo1.Text & "'", Con, adOpenUnspecified, adLockReadOnly

While Not rec.EOF
Text6.Text = rec(0)
rec.MoveNext
Wend
Con.Close
End Sub

'Selektimi i te dhenave nga tabela Specializimi dhe vendosja e te dhenave tek Combobox
Private Sub Combo2_Add()
Combo2.Clear
Dim strng As String
Con = "Provider=MSDASQL.1;Data Source=Juli;Initial Catalog=SQLEXPRESS"
Con.Open "Driver=SQL Server;Server=.\SQLEXPRESS;Database=Hospital;uid=;pwd=;"
rec.Open "select llojSpecializimi from Specializimi", Con, adOpenUnspecified, adLockReadOnly
Combo2.AddItem "Selekto"
Combo2.Text = Me.Combo2.List(0)
While Not rec.EOF
Combo2.AddItem rec(0)
rec.MoveNext
Wend
Con.Close

End Sub

Private Sub Combo2_Click()
Dim strng As String
Con = "Provider=MSDASQL.1;Data Source=Juli;Initial Catalog=SQLEXPRESS"
Con.Open "Driver=SQL Server;Server=.\SQLEXPRESS;Database=Hospital;uid=;pwd=;"
rec.Open "select idSpecializimi from Specializimi Where llojSpecializimi = '" & Combo2.Text & "'", Con, adOpenUnspecified, adLockReadOnly

While Not rec.EOF
Text9.Text = rec(0)
rec.MoveNext
Wend
Con.Close

End Sub

'Selektimi i te dhenave nga tabela Departamenti dhe vendosja e te dhenave tek Combobox

Private Sub Combo3_Add()
Combo3.Clear
Dim strng As String
Con = "Provider=MSDASQL.1;Data Source=Juli;Initial Catalog=SQLEXPRESS"
Con.Open "Driver=SQL Server;Server=.\SQLEXPRESS;Database=Hospital;uid=;pwd=;"
rec.Open "select EmriDepartamentit from Departamenti", Con, adOpenUnspecified, adLockReadOnly
Combo3.AddItem "Selekto"
Combo3.Text = Me.Combo3.List(0)
While Not rec.EOF
Combo3.AddItem rec(0)
rec.MoveNext
Wend
Con.Close

End Sub
Private Sub Combo3_Click()
Dim strng As String
Con = "Provider=MSDASQL.1;Data Source=Juli;Initial Catalog=SQLEXPRESS"
Con.Open "Driver=SQL Server;Server=.\SQLEXPRESS;Database=Hospital;uid=;pwd=;"
rec.Open "select IdDepartamenti from Departamenti Where EmriDepartamentit = '" & Combo3.Text & "'", Con, adOpenUnspecified, adLockReadOnly

While Not rec.EOF
Text10.Text = rec(0)
rec.MoveNext
Wend
Con.Close

End Sub

Private Sub Command49_Click()
Text1.Text = RndWord(7)
End Sub

'Klikimi i butonit I RI dhe fshirja e fushave ne gjendje fillestare

Private Sub Command1_Click()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text7.Text = ""
Text8.Text = ""
Combo1.Text = Me.Combo1.List(0)
Combo2.Text = Me.Combo2.List(0)
Combo3.Text = Me.Combo3.List(0)
DTPicker1.Value = Format(Now(), "dd/MM/yyyy")
DTPicker2.Value = Format(Now(), "dd/MM/yyyy")
End Sub

Private Sub Command2_Click()
roli = "2"
status1 = "Aktiv"
Con = "Provider=MSDASQL.1;Data Source=Juli;Initial Catalog=SQLEXPRESS"
Con.Open "Driver=SQL Server;Server=.\SQLEXPRESS;Database=Hospital;uid=;pwd=;"
If Text1.Text = "" Or Text2.Text = "" Or Text3.Text = "" Or Text4.Text = "" Or Text5.Text = "" Or Text6.Text = "" Or Text7.Text = "" Or Text8.Text = "" Or Text9.Text = "" Or Text10.Text = "" Or Text15.Text = "" Or Text16.Text = "" Then
MsgBox "PLOTESONI FUSHAT ! "
Else
Set rs = Con.Execute( _
        "INSERT INTO Mjeku(IdMjek, Emri, Atesia, Mbiemri,Gjinia, Datelindja, idKualifikimi, Kontakt, Email, idSpecializimi, IdDepartamenti, Dt_Punesimit,StatusPune)VALUES('" & Text1.Text & "','" & Text2.Text & "','" & Text3.Text & "','" & Text4.Text & "','" & Text5.Text & "', '" & Format$(DTPicker1.Value, "yyyy.mm.dd") & "' ,'" & Text6.Text & "','" & Text7.Text & "','" & Text8.Text & "','" & Text9.Text & "','" & Text10.Text & "', '" & Format$(DTPicker2.Value, "yyyy.mm.dd") & "', '" & status1 & "')")
Set rs = Con.Execute( _
        "INSERT INTO Login(IdRoli,IdMjeku,Perdoruesi,Fjalekalimi)Values('" & roli & "', '" & Text1.Text & "','" & Text15.Text & "', '" & Text16.Text & "' )")
        
        
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

Private Sub Command3_Click()
Admin.SSTab2.Tab = 1
End Sub

Private Sub Command4_Click()
Unload Me
End Sub



'Selektimi i gjinise me OPTION button

Private Sub Option1_Click()
Text5.Text = "M"
End Sub

Private Sub Option2_Click()
Text5.Text = "F"
End Sub
'==================================================FUND SHTIMI I MJEKUT=====================================================================


'==================================================NDRYSHIMI I TE DHENAVE TE MJEKUT=========================================================

'Klikimi i butonit I RI dhe kalimi ne tabin SHTO MJEK

Private Sub Command5_Click()
Admin.SSTab2.Tab = 0
End Sub

' Ruajtja e te dhenave te modifikuara

Private Sub Command6_Click()
Con = "Provider=MSDASQL.1;Data Source=Juli;Initial Catalog=SQLEXPRESS"
Con.Open "Driver=SQL Server;Server=.\SQLEXPRESS;Database=Hospital;uid=;pwd=;"

If Text11.Text = "" Or Text12.Text = "" Or Text13.Text = "" Or Text14.Text = "" Or Combo18 = "" Or DTPicker3.Value = "" Or Text17.Text = "" Or Text18.Text = "" Or Text20.Text = "" Or DTPicker4.Value = "" Or Text22.Text = "" Or Combo6.Text = "" Then
MsgBox "PLOTESO FUSHAT."
Else
    On Error GoTo ProcError
    
Set rs = Con.Execute("Update Mjeku Set Emri = '" & Text12.Text & "', Atesia = '" & Text13.Text & "', Mbiemri = '" & Text14.Text & "', Gjinia = '" & Combo18.Text & "', Datelindja = '" & Format$(DTPicker3.Value, "yyyy.mm.dd") & "', idKualifikimi = '" & Text17.Text & "', Kontakt = '" & Text18.Text & "', Email = '" & Text19.Text & "', idSpecializimi  = '" & Text20.Text & "', IdDepartamenti  = '" & Text22.Text & "',  Dt_Punesimit = '" & Format$(DTPicker4.Value, "yyyy.mm.dd") & "', StatusPune = '" & Combo9 & "' Where IdMjek = '" & Text11.Text & "' ")
    Set rs = Con.Execute("Update Login Set Perdoruesi = '" & Text21.Text & "', Fjalekalimi = '" & Text33.Text & "' Where IdMjeku = '" & Text11.Text & "' ")

 
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
' Mbushja e COMBOBOX-eve
Private Sub Combo6_Add()
Dim strng As String
Con = "Provider=MSDASQL.1;Data Source=Juli;Initial Catalog=SQLEXPRESS"
Con.Open "Driver=SQL Server;Server=.\SQLEXPRESS;Database=Hospital;uid=;pwd=;"
rec.Open "select llojKualifikim from Kualifikimi", Con, adOpenUnspecified, adLockReadOnly
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
rec.Open "select idKualifikim from Kualifikimi Where llojKualifikim = '" & Combo6 & "'", Con, adOpenUnspecified, adLockReadOnly
While Not rec.EOF
Text17.Text = rec(0)
rec.MoveNext
Wend
Con.Close
End Sub
Sub Combo7_Add()

Dim strng As String
Con = "Provider=MSDASQL.1;Data Source=Juli;Initial Catalog=SQLEXPRESS"
Con.Open "Driver=SQL Server;Server=.\SQLEXPRESS;Database=Hospital;uid=;pwd=;"
rec.Open "select llojSpecializimi from Specializimi", Con, adOpenUnspecified, adLockReadOnly
Combo7.AddItem "Selekto"
Combo7.Text = Me.Combo7.List(0)
While Not rec.EOF
Combo7.AddItem rec(0)
rec.MoveNext
Wend
Con.Close
End Sub

Private Sub Combo7_Click()

Con = "Provider=MSDASQL.1;Data Source=Juli;Initial Catalog=SQLEXPRESS"
Con.Open "Driver=SQL Server;Server=.\SQLEXPRESS;Database=Hospital;uid=;pwd=;"
rec.Open "select idSpecializimi from Specializimi Where llojSpecializimi = '" & Combo7 & "'", Con, adOpenUnspecified, adLockReadOnly

While Not rec.EOF
Text20.Text = rec(0)
rec.MoveNext
Wend
Con.Close
End Sub

Private Sub Combo8_Add()
Dim strng As String
Con = "Provider=MSDASQL.1;Data Source=Juli;Initial Catalog=SQLEXPRESS"
Con.Open "Driver=SQL Server;Server=.\SQLEXPRESS;Database=Hospital;uid=;pwd=;"
rec.Open "select EmriDepartamentit from Departamenti", Con, adOpenUnspecified, adLockReadOnly
Combo8.AddItem "Selekto"
Combo8.Text = Me.Combo8.List(0)
While Not rec.EOF
Combo8.AddItem rec(0)
rec.MoveNext
Wend
Con.Close

End Sub
Private Sub Combo8_Click()
Con = "Provider=MSDASQL.1;Data Source=Juli;Initial Catalog=SQLEXPRESS"
Con.Open "Driver=SQL Server;Server=.\SQLEXPRESS;Database=Hospital;uid=;pwd=;"
rec.Open "select idDepartamenti from Departamenti Where EmriDepartamentit = '" & Combo8 & "'", Con, adOpenUnspecified, adLockReadOnly

While Not rec.EOF
Text22.Text = rec(0)
rec.MoveNext
Wend
Con.Close

End Sub
' Klikimi i butonit RIRUAJ dhe modifikimi i te dhenave

Private Sub Command7_Click()
status = "Pasiv"
Con = "Provider=MSDASQL.1;Data Source=Juli;Initial Catalog=SQLEXPRESS"
Con.Open "Driver=SQL Server;Server=.\SQLEXPRESS;Database=Hospital;uid=;pwd=;"
If Text11.Text = "" Or Text12.Text = "" Or Text13.Text = "" Or Text14.Text = "" Or Combo18.Text = "" Or Text16.Text = "" Or Text17.Text = "" Or Text18.Text = "" Or Text19.Text = "" Or Text20.Text = "" Or Text21.Text = "" Or Text22.Text = "" Or Combo6.Text = "" Then
MsgBox "PLOTESO FUSHAT."
Else
    On Error GoTo ProcError
Set rs = Con.Execute("Update Mjeku Set StatusPune = '" & status & "' Where IdMjek = '" & Text11.Text & "' ")
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

Private Sub Command8_Click()
Combo4 = " "
Combo5 = " "
Text11.Text = " "
Text12.Text = " "
Text13.Text = " "
Text14.Text = " "
Combo18 = " "
DTPicker3.Value = Format(Now(), "dd/MM/yyyy")
Combo6 = " "
Text17.Text = " "
Text18.Text = " "
Text19.Text = " "
Combo7 = " "
Text20.Text = " "
Combo8 = " "
Text22.Text = " "
DTPicker4.Value = Format(Now(), "dd/MM/yyyy")
Combo9 = " "
Text21.Text = " "
Text33.Text = " "

Adodc1.RecordSource = "Select IdMjek, Emri, Atesia, Mbiemri, Gjinia, Datelindja, llojKualifikim, Kontakt, Email, llojSpecializimi, EmriDepartamentit, Dt_Punesimit,StatusPune,Perdoruesi,Fjalekalimi from Mjeku, Departamenti,Kualifikimi, Specializimi, Login  Where Departamenti.IdDepartamenti = Mjeku.IdDepartamenti And Kualifikimi.idKualifikim = Mjeku.idKualifikimi And Specializimi.idSpecializimi =  Mjeku.idSpecializimi And Login.IdMjeku = Mjeku.IdMjek"
Adodc1.Refresh
DataGrid1.Refresh

End Sub

Private Sub Command9_Click()
Unload Me
End Sub

' Kerkimi me ane te Emrit dhe Mbiemrit te Mjekut

Sub Combo4_Add()
Combo4.Clear

Dim strng As String


Con = "Provider=MSDASQL.1;Data Source=Juli;Initial Catalog=SQLEXPRESS"
Con.Open "Driver=SQL Server;Server=.\SQLEXPRESS;Database=Hospital;uid=;pwd=;"
rec.Open "select Emri from Mjeku", Con, adOpenUnspecified, adLockReadOnly

Combo4.AddItem "Selekto"
Combo4.Text = Me.Combo4.List(0)

While Not rec.EOF
Combo4.AddItem rec(0)
rec.MoveNext
Wend
Con.Close
End Sub

Private Sub Combo4_Click()
If Combo4.ListIndex = 0 Then
        Adodc1.RecordSource = "select IdMjek, Emri, Atesia, Mbiemri, Gjinia, KOntakt, Email, Dt_Punesimit,EmriDepartamentit, Datelindja, llojKualifikim,llojSpecializimi from Mjeku, Departamenti,Kualifikimi, Specializimi Where Departamenti.IdDepartamenti = Mjeku.IdDepartamenti And Kualifikimi.idKualifikim = Mjeku.idKualifikimi And Specializimi.idSpecializimi =  Mjeku.idSpecializimi"
    Else
        Adodc1.RecordSource = "select IdMjek, Emri, Atesia, Mbiemri, Gjinia, KOntakt, Email, Dt_Punesimit,EmriDepartamentit, Datelindja, llojKualifikim,llojSpecializimi from Mjeku, Departamenti,Kualifikimi, Specializimi Where Departamenti.IdDepartamenti = Mjeku.IdDepartamenti And Kualifikimi.idKualifikim = Mjeku.idKualifikimi And Specializimi.idSpecializimi =  Mjeku.idSpecializimi And Emri = '" & Combo4 & "'"
    End If
    
    Combo4.Refresh
    Adodc1.Refresh
DataGrid1.Refresh
Combo5.Enabled = True

Con = "Provider=MSDASQL.1;Data Source=Juli;Initial Catalog=SQLEXPRESS"
Con.Open "Driver=SQL Server;Server=.\SQLEXPRESS;Database=Hospital;uid=;pwd=;"
rec.Open "select Mbiemri from Mjeku Where Emri = '" & Combo4 & "'", Con, adOpenUnspecified, adLockReadOnly

Combo5.AddItem "Selekto"
Combo5.Text = Me.Combo5.List(0)

While Not rec.EOF
Combo5.AddItem rec(0)
rec.MoveNext
Wend
Con.Close


End Sub
Private Sub Combo5_Click()
If Combo5.ListIndex = 0 Then
        Adodc1.RecordSource = "select IdMjek, Emri, Atesia, Mbiemri, Gjinia, KOntakt, Email, Dt_Punesimit,EmriDepartamentit, Datelindja, llojKualifikim,llojSpecializimi from Mjeku, Departamenti,Kualifikimi, Specializimi Where Departamenti.IdDepartamenti = Mjeku.IdDepartamenti And Kualifikimi.idKualifikim = Mjeku.idKualifikimi And Specializimi.idSpecializimi =  Mjeku.idSpecializimi"
    Else
        Adodc1.RecordSource = "select IdMjek, Emri, Atesia, Mbiemri, Gjinia, KOntakt, Email, Dt_Punesimit,EmriDepartamentit, Datelindja, llojKualifikim,llojSpecializimi from Mjeku, Departamenti,Kualifikimi, Specializimi Where Departamenti.IdDepartamenti = Mjeku.IdDepartamenti And Kualifikimi.idKualifikim = Mjeku.idKualifikimi And Specializimi.idSpecializimi =  Mjeku.idSpecializimi And Emri = '" & Combo4 & "' And Mbiemri = '" & Combo5 & "'"
    End If
    
    Combo5.Refresh
    Adodc1.Refresh
DataGrid1.Refresh
End Sub


'Mbushja e te dhenave nga DataGrid

Private Sub DataGrid1_Click()
DataGrid1.Col = 0
Admin.Text11.Text = Admin.DataGrid1.Text
DataGrid1.Col = 1
Admin.Text12.Text = Admin.DataGrid1.Text
Admin.Combo4.Text = Admin.DataGrid1.Text
DataGrid1.Col = 2
Admin.Text13.Text = Admin.DataGrid1.Text
DataGrid1.Col = 3
Admin.Text14.Text = Admin.DataGrid1.Text
Admin.Combo5.Text = Admin.DataGrid1.Text
DataGrid1.Col = 4
Admin.Combo18.Text = Admin.DataGrid1.Text
DataGrid1.Col = 5
Admin.DTPicker3.Value = Admin.DataGrid1.Text
DataGrid1.Col = 6
Admin.Combo6.Text = Admin.DataGrid1.Text
Con = "Provider=MSDASQL.1;Data Source=Juli;Initial Catalog=SQLEXPRESS"
Con.Open "Driver=SQL Server;Server=.\SQLEXPRESS;Database=Hospital;uid=;pwd=;"
rec.Open "select idKualifikim from Kualifikimi Where llojKualifikim = '" & Combo6 & "'", Con, adOpenUnspecified, adLockReadOnly
While Not rec.EOF
Text17.Text = rec(0)
rec.MoveNext
Wend
Con.Close
DataGrid1.Col = 7
Admin.Text18.Text = Admin.DataGrid1.Text
DataGrid1.Col = 8
Admin.Text19.Text = Admin.DataGrid1.Text
DataGrid1.Col = 9
Admin.Combo7.Text = Admin.DataGrid1.Text
Con = "Provider=MSDASQL.1;Data Source=Juli;Initial Catalog=SQLEXPRESS"
Con.Open "Driver=SQL Server;Server=.\SQLEXPRESS;Database=Hospital;uid=;pwd=;"
rec.Open "select idSpecializimi from Specializimi Where llojSpecializimi = '" & Combo7 & "'", Con, adOpenUnspecified, adLockReadOnly
While Not rec.EOF
Text20.Text = rec(0)
rec.MoveNext
Wend
Con.Close
DataGrid1.Col = 10
Admin.Combo8.Text = Admin.DataGrid1.Text
Con = "Provider=MSDASQL.1;Data Source=Juli;Initial Catalog=SQLEXPRESS"
Con.Open "Driver=SQL Server;Server=.\SQLEXPRESS;Database=Hospital;uid=;pwd=;"
rec.Open "select idDepartamenti from Departamenti Where EmriDepartamentit = '" & Combo8 & "'", Con, adOpenUnspecified, adLockReadOnly
While Not rec.EOF
Text22.Text = rec(0)
rec.MoveNext
Wend
Con.Close
DataGrid1.Col = 11
Admin.DTPicker4.Value = Admin.DataGrid1.Text

DataGrid1.Col = 12
Admin.Combo9.Text = Admin.DataGrid1.Text

DataGrid1.Col = 13
Admin.Text21.Text = Admin.DataGrid1.Text
DataGrid1.Col = 14
Admin.Text33.Text = Admin.DataGrid1.Text

End Sub

'Gjenerimi i Raportit

Private Sub Command44_Click()
MjekuRaport.Show
End Sub

'==============================================FUND NDRYSHIMI I MJEKUT==========================================================================

'=============================================SHTIMI DHE MODIFIKIMI I DEPARTAMENTIT=============================================================
Private Sub Command14_Click()
Admin.SSTab3.Tab = 0
End Sub
Private Sub Command10_Click()
Text23.Text = ""
Text24.Text = ""
End Sub

' Shtimi i Departamentit
Private Sub Command11_Click()
Con = "Provider=MSDASQL.1;Data Source=Juli;Initial Catalog=SQLEXPRESS"
Con.Open "Driver=SQL Server;Server=.\SQLEXPRESS;Database=Hospital;uid=;pwd=;"
If Text23.Text = "" Or Text24.Text = "" Then
MsgBox "PLOTESONI FUSHAT"
Else
On Error GoTo ProcError
Set rs = Con.Execute( _
        "INSERT INTO Departamenti(IdDepartamenti, EmriDepartamentit )VALUES('" & Text23.Text & "','" & Text24.Text & "' )")
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
Private Sub Command12_Click()
Admin.SSTab3.Tab = 1
End Sub
Private Sub Command13_Click()
Unload Me
End Sub



Sub Combo10_Add()
Combo10.Clear
Dim strng As String


Con = "Provider=MSDASQL.1;Data Source=Juli;Initial Catalog=SQLEXPRESS"
Con.Open "Driver=SQL Server;Server=.\SQLEXPRESS;Database=Hospital;uid=;pwd=;"
rec.Open "select EmriDepartamentit from Departamenti", Con, adOpenUnspecified, adLockReadOnly



Combo10.AddItem "Selekto"
Combo10.Text = Me.Combo10.List(0)
While Not rec.EOF
Combo10.AddItem rec(0)
rec.MoveNext
Wend
Con.Close
End Sub


Private Sub Combo10_Click()
If Combo10.ListIndex = 0 Then
        Adodc2.RecordSource = "select * from Departamenti "
    Else
        Adodc2.RecordSource = "select * from Departamenti where EmriDepartamentit = '" & Combo10 & "'"
    End If
    
    Combo10.Refresh
    Adodc2.Refresh
DataGrid2.Refresh

End Sub

Private Sub DataGrid2_Click()
DataGrid2.Col = 0
Admin.Text25.Text = Admin.DataGrid2.Text
DataGrid2.Col = 1
Admin.Text26.Text = Admin.DataGrid2.Text
Admin.Combo10.Text = Admin.DataGrid2.Text
End Sub
Private Sub Command16_Click()
Text25.Text = " "
Text26.Text = " "
Combo10.Text = Me.Combo10.List(0)
Adodc2.RecordSource = "select * from Departamenti "
Adodc2.Refresh
DataGrid2.Refresh
End Sub
Private Sub Command15_Click()
Con = "Provider=MSDASQL.1;Data Source=Juli;Initial Catalog=SQLEXPRESS"
Con.Open "Driver=SQL Server;Server=.\SQLEXPRESS;Database=Hospital;uid=;pwd=;"

If Text25.Text = "" Or Text26.Text = "" Then
MsgBox "PLOTESO FUSHAT."
Else
    On Error GoTo ProcError
    
Set rs = Con.Execute("Update Departamenti Set EmriDepartamentit = '" & Text26.Text & "' Where IdDepartamenti = '" & Text25.Text & "' ")
    
 
MsgBox "MODIFIKIMI I TE DHENAVE U KRYE ME SUKSES."

End If
Con.Close
ProcExit:
Exit Sub
ProcError:


'MsgBox "Error Number: " & Err.Number & " With The Description ->> " & Err.Description & " <<- Occured."

MsgBox "TE DHENAT NUK U MODIFIKUAN."
Con.Close
Resume ProcExit
Con.Close

End Sub

Private Sub Command17_Click()
Unload Me
End Sub

'================================================FUND DEPARTAMENTI===================================================================================

'=============================================== ILACE / ANALIZA / INJEKSIONE =======================================================================
     '======================================================== ILACET  ==============================================================================
Private Sub Command18_Click()
Text27.Text = ""
Text28.Text = ""
Text29.Text = ""
End Sub

Private Sub Command19_Click()
Dim gjendja As String
gjendja = "Ka gjendje"

Con = "Provider=MSDASQL.1;Data Source=Juli;Initial Catalog=SQLEXPRESS"
Con.Open "Driver=SQL Server;Server=.\SQLEXPRESS;Database=Hospital;uid=;pwd=;"


If Text27.Text = "" Or Text28.Text = "" Or Text29.Text = "" Then
MsgBox "PLOTESONI FUSHAT"
Else

On Error GoTo ProcError
Set rs = Con.Execute( _
        "INSERT INTO Ilace(IdIlace, EmerIlace, CmimiIlace, Gjendje)VALUES('" & Text27.Text & "','" & Text28.Text & "', '" & Text29.Text & "', '" & gjendja & "' )")

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
Private Sub Command20_Click()
Admin.SSTab5.Tab = 1
End Sub
Private Sub Command22_Click()
Admin.SSTab5.Tab = 0
End Sub
Private Sub Command21_Click()
Unload Me
End Sub
Private Sub Combo11_Add()
Combo11.Clear
Dim strng As String
Con = "Provider=MSDASQL.1;Data Source=Juli;Initial Catalog=SQLEXPRESS"
Con.Open "Driver=SQL Server;Server=.\SQLEXPRESS;Database=Hospital;uid=;pwd=;"
rec.Open "select EmerIlace from Ilace", Con, adOpenUnspecified, adLockReadOnly
Combo11.AddItem "Selekto"
Combo11.Text = Me.Combo11.List(0)
While Not rec.EOF
Combo11.AddItem rec(0)
rec.MoveNext
Wend
Con.Close
End Sub
Private Sub Combo11_Click()
If Combo11.ListIndex = 0 Then
        Adodc3.RecordSource = "select * from Ilace "
    Else
        Adodc3.RecordSource = "select * from Ilace where EmerIlace = '" & Combo11 & "'"
    End If
    
    Combo11.Refresh
    Adodc3.Refresh
DataGrid3.Refresh

End Sub
Private Sub Command23_Click()

Con = "Provider=MSDASQL.1;Data Source=Juli;Initial Catalog=SQLEXPRESS"
Con.Open "Driver=SQL Server;Server=.\SQLEXPRESS;Database=Hospital;uid=;pwd=;"

If Text30.Text = "" Or Text31.Text = "" Or Text32.Text = "" Or Combo12 = "" Then
MsgBox "PLOTESO FUSHAT."
Else
    On Error GoTo ProcError
    
Set rs = Con.Execute("Update Ilace Set EmerIlace = '" & Text31.Text & "', CmimiIlace = '" & Text32.Text & "', Gjendje = '" & Combo12 & "' Where IdIlace = '" & Text30.Text & "' ")
    
 
MsgBox "MODIFIKIMI I TE DHENAVE U KRYE ME SUKSES."

End If
Con.Close
ProcExit:
Exit Sub
ProcError:


'MsgBox "Error Number: " & Err.Number & " With The Description ->> " & Err.Description & " <<- Occured."

MsgBox "TE DHENAT NUK U MODIFIKUAN."
Con.Close
Resume ProcExit
Con.Close

End Sub
Private Sub Command24_Click()
gjendjeIlac = "Ska Gjendje"
Con = "Provider=MSDASQL.1;Data Source=Juli;Initial Catalog=SQLEXPRESS"
Con.Open "Driver=SQL Server;Server=.\SQLEXPRESS;Database=Hospital;uid=;pwd=;"
If Text30.Text = "" Or Text31.Text = "" Or Text32.Text = "" Or Combo12.Text = "" Then
MsgBox "PLOTESO FUSHAT."
Else
    On Error GoTo ProcError
Set rs = Con.Execute("Update Ilace Set Gjendje = '" & gjendjeIlac & "' Where idIlace = '" & Text30.Text & "' ")
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

Private Sub Command25_Click()
Text30.Text = " "
Text31.Text = " "
Text32.Text = " "
Combo11.Text = "Selekto"
Combo12.Text = " "
Adodc3.RecordSource = "select * from Ilace "
Adodc3.Refresh
DataGrid3.Refresh
End Sub
Private Sub DataGrid3_Click()
DataGrid3.Col = 0
Admin.Text30 = Admin.DataGrid3.Text
DataGrid3.Col = 1
Admin.Text31 = Admin.DataGrid3.Text
Admin.Combo11 = Admin.DataGrid3.Text
DataGrid3.Col = 2
Admin.Text32 = Admin.DataGrid3.Text
DataGrid3.Col = 3
Admin.Combo12 = Admin.DataGrid3.Text

End Sub

'Gjenerimi i Raportit

Private Sub Command45_Click()
IlaceRaport.Show
End Sub
     '======================================================== FUND ILACET  ==============================================================================
     '======================================================== ANALIZAT ==================================================================================

Private Sub Command27_Click()
Text33.Text = " "
Text34.Text = " "
Text35.Text = " "
End Sub
Private Sub Command28_Click()
Con = "Provider=MSDASQL.1;Data Source=Juli;Initial Catalog=SQLEXPRESS"
Con.Open "Driver=SQL Server;Server=.\SQLEXPRESS;Database=Hospital;uid=;pwd=;"


If Text33.Text = "" Or Text34.Text = "" Or Text35.Text = "" Then
MsgBox "PLOTESONI FUSHAT"
Else

On Error GoTo ProcError
Set rs = Con.Execute( _
        "INSERT INTO Analiza(EmerAnaliza, Kosto)VALUES('" & Text34.Text & "', '" & Text35.Text & "' )")

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

Private Sub Command29_Click()
Admin.SSTab6.Tab = 1
End Sub

Private Sub Command30_Click()
Unload Me
End Sub
Private Sub Command31_Click()
Admin.SSTab6.Tab = 0
End Sub
Private Sub Command32_Click()
Con = "Provider=MSDASQL.1;Data Source=Juli;Initial Catalog=SQLEXPRESS"
Con.Open "Driver=SQL Server;Server=.\SQLEXPRESS;Database=Hospital;uid=;pwd=;"

If Text36.Text = "" Or Text37.Text = "" Or Text38.Text = "" Then
MsgBox "PLOTESO FUSHAT."
Else
    On Error GoTo ProcError
    
Set rs = Con.Execute("Update Analiza Set EmerAnaliza = '" & Text37.Text & "', Kosto = '" & Text38.Text & "' Where IdAnaliza = '" & Text36.Text & "' ")
 
MsgBox "MODIFIKIMI I TE DHENAVE U KRYE ME SUKSES."

End If
Con.Close
ProcExit:
Exit Sub
ProcError:


'MsgBox "Error Number: " & Err.Number & " With The Description ->> " & Err.Description & " <<- Occured."

MsgBox "TE DHENAT NUK U MODIFIKUAN."
Con.Close
Resume ProcExit
Con.Close


End Sub
Private Sub Command33_Click()
Text36.Text = " "
Text37.Text = " "
Text38.Text = " "

Combo13.Text = Me.Combo13.List(0)
Adodc4.RecordSource = "select * from Analiza "
Adodc4.Refresh
DataGrid4.Refresh
End Sub

Private Sub Combo13_Add()
Combo13.Clear
Dim strng As String
Con = "Provider=MSDASQL.1;Data Source=Juli;Initial Catalog=SQLEXPRESS"
Con.Open "Driver=SQL Server;Server=.\SQLEXPRESS;Database=Hospital;uid=;pwd=;"
rec.Open "select EmerAnaliza from Analiza", Con, adOpenUnspecified, adLockReadOnly
Combo13.AddItem "Selekto"
Combo13.Text = Me.Combo13.List(0)
While Not rec.EOF
Combo13.AddItem rec(0)
rec.MoveNext
Wend
Con.Close
End Sub

Private Sub Combo13_Click()
If Combo13.ListIndex = 0 Then
        Adodc4.RecordSource = "select * from Ilace "
    Else
        Adodc4.RecordSource = "select * from Analiza where EmerAnaliza = '" & Combo13 & "'"
    End If
    
    Combo13.Refresh
    Adodc4.Refresh
DataGrid4.Refresh
End Sub
Private Sub DataGrid4_Click()
DataGrid4.Col = 0
Admin.Text36.Text = Admin.DataGrid4.Text
DataGrid4.Col = 1
Admin.Text37.Text = Admin.DataGrid4.Text
Admin.Combo13.Text = Admin.DataGrid4.Text
DataGrid4.Col = 2
Admin.Text38.Text = Admin.DataGrid4.Text

End Sub

'Gjenerimi i raportit
Private Sub Command46_Click()
AnalizaRaport.Show
End Sub
     
     '======================================================== FUND ANALIZAT ==================================================================================

     '======================================================== INJEKSIONET ==================================================================================
Private Sub Command35_Click()
Text39.Text = " "
Text40.Text = " "
Text41.Text = " "
End Sub

Private Sub Command36_Click()
InjeksionGjendje = "Ka Gjendje"
Con = "Provider=MSDASQL.1;Data Source=Juli;Initial Catalog=SQLEXPRESS"
Con.Open "Driver=SQL Server;Server=.\SQLEXPRESS;Database=Hospital;uid=;pwd=;"


If Text39.Text = " " Or Text40.Text = " " Or Text41.Text = " " Or Combo17 = " " Or Text46.Text = " " Then
MsgBox "PLOTESONI FUSHAT"
Else

On Error GoTo ProcError
Set rs = Con.Execute( _
        "INSERT INTO Injeksionet(IdInjeksion, EmerInjeksion, IdTipi, CmimInjeksion, Gjendje)VALUES('" & Text39.Text & "','" & Text40.Text & "', '" & Text46.Text & "', '" & Text41.Text & "', '" & InjeksionGjendje & "' )")

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
Private Sub Combo17_Add()
Dim strng As String

Con = "Provider=MSDASQL.1;Data Source=Juli;Initial Catalog=SQLEXPRESS"
Con.Open "Driver=SQL Server;Server=.\SQLEXPRESS;Database=Hospital;uid=;pwd=;"
rec.Open "select TipiInjeksionit from TipInjeksioni", Con, adOpenUnspecified, adLockReadOnly

Combo17.AddItem "Selekto"
Combo17.Text = Me.Combo17.List(0)
While Not rec.EOF
Combo17.AddItem rec(0)
rec.MoveNext
Wend
Con.Close
End Sub

Private Sub Combo17_Click()
Con = "Provider=MSDASQL.1;Data Source=Juli;Initial Catalog=SQLEXPRESS"
Con.Open "Driver=SQL Server;Server=.\SQLEXPRESS;Database=Hospital;uid=;pwd=;"
rec.Open "select IdTipInjeksioni from TipInjeksioni Where TipiInjeksionit = '" & Combo17 & "'", Con, adOpenUnspecified, adLockReadOnly



While Not rec.EOF
Text46.Text = rec(0)
rec.MoveNext
Wend
Con.Close
End Sub

Private Sub Command37_Click()
Admin.SSTab7.Tab = 1
End Sub

Private Sub Command38_Click()
Unload Me
End Sub

Private Sub Command39_Click()
Admin.SSTab7.Tab = 0
End Sub

Private Sub Command40_Click()
Con = "Provider=MSDASQL.1;Data Source=Juli;Initial Catalog=SQLEXPRESS"
Con.Open "Driver=SQL Server;Server=.\SQLEXPRESS;Database=Hospital;uid=;pwd=;"

If Text42.Text = "" Or Text43.Text = "" Or Text44.Text = "" Or Combo15.Text = "" Or Combo16.Text = "" Or Text45.Text = "" Then
MsgBox "PLOTESO FUSHAT."
Else
    On Error GoTo ProcError
    
Set rs = Con.Execute("Update Injeksionet Set EmerInjeksion = '" & Text43.Text & "', IdTipi = '" & Text45.Text & "', CmimInjeksion = '" & Text44.Text & "', Gjendje = '" & Combo16.Text & "' Where IdInjeksion = '" & Text42.Text & "' ")
    
 
MsgBox "MODIFIKIMI I TE DHENAVE U KRYE ME SUKSES."

End If
Con.Close
ProcExit:
Exit Sub
ProcError:


'MsgBox "Error Number: " & Err.Number & " With The Description ->> " & Err.Description & " <<- Occured."

MsgBox "TE DHENAT NUK U MODIFIKUAN."
Con.Close
Resume ProcExit
Con.Close

End Sub
Private Sub Combo15_Add()
Dim strng As String
Con = "Provider=MSDASQL.1;Data Source=Juli;Initial Catalog=SQLEXPRESS"
Con.Open "Driver=SQL Server;Server=.\SQLEXPRESS;Database=Hospital;uid=;pwd=;"
rec.Open "select TipiInjeksionit from TipInjeksioni", Con, adOpenUnspecified, adLockReadOnly

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
rec.Open "select IdTipInjeksioni from TipInjeksioni Where TipiInjeksionit = '" & Combo15 & "'", Con, adOpenUnspecified, adLockReadOnly
While Not rec.EOF
Text45.Text = rec(0)
rec.MoveNext
Wend
Con.Close
End Sub

Private Sub DataGrid5_Click()
DataGrid5.Col = 0
Admin.Text42.Text = Admin.DataGrid5.Text

DataGrid5.Col = 1
Admin.Text43.Text = Admin.DataGrid5.Text
Admin.Combo14.Text = Admin.DataGrid5.Text

DataGrid5.Col = 2

Admin.Combo15.Text = Admin.DataGrid5.Text

Con = "Provider=MSDASQL.1;Data Source=Juli;Initial Catalog=SQLEXPRESS"
Con.Open "Driver=SQL Server;Server=.\SQLEXPRESS;Database=Hospital;uid=;pwd=;"
rec.Open "Select IdTipInjeksioni From TipInjeksioni Where TipInjeksioni.TipiInjeksionit = '" & Combo15 & "'", Con, adOpenUnspecified, adLockReadOnly
While Not rec.EOF
Text45.Text = rec(0)
rec.MoveNext
Wend
Con.Close

DataGrid5.Col = 3

Admin.Text44.Text = Admin.DataGrid5.Text

DataGrid5.Col = 4

Admin.Combo16.Text = Admin.DataGrid5.Text

End Sub
Private Sub Command41_Click()
GjendjeInjeksioni = "Ska Gjendje"
Con = "Provider=MSDASQL.1;Data Source=Juli;Initial Catalog=SQLEXPRESS"
Con.Open "Driver=SQL Server;Server=.\SQLEXPRESS;Database=Hospital;uid=;pwd=;"
If Text42.Text = "" Or Text43.Text = "" Or Text44.Text = "" Or Combo16.Text = "" Then
MsgBox "PLOTESO FUSHAT."
Else
    On Error GoTo ProcError
Set rs = Con.Execute("Update Injeksionet Set Gjendje = '" & GjendjeInjeksioni & "' Where IdInjeksion = '" & Text42.Text & "' ")
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

Private Sub Command42_Click()
Text42.Text = " "
Text43.Text = " "
Text44.Text = " "
Combo14 = "Selekto"
Combo15 = " "
Text45.Text = " "
Combo16 = " "
Adodc5.RecordSource = "Select IdInjeksion, EmerInjeksion, TipiInjeksionit, CmimInjeksion, Gjendje From Injeksionet, TipInjeksioni Where TipInjeksioni.IdTipInjeksioni = Injeksionet.IdTipi "
Adodc5.Refresh
DataGrid5.Refresh
End Sub
Private Sub Combo14_Add()
Combo14.Clear
Dim strng As String
Con = "Provider=MSDASQL.1;Data Source=Juli;Initial Catalog=SQLEXPRESS"
Con.Open "Driver=SQL Server;Server=.\SQLEXPRESS;Database=Hospital;uid=;pwd=;"
rec.Open "select EmerInjeksion from Injeksionet", Con, adOpenUnspecified, adLockReadOnly
Combo14.AddItem "Selekto"
Combo14.Text = Me.Combo14.List(0)
While Not rec.EOF
Combo14.AddItem rec(0)
rec.MoveNext
Wend
Con.Close
End Sub

Private Sub Combo14_Click()
If Combo14.ListIndex = 0 Then
        Adodc5.RecordSource = "select IdInjeksion, EmerInjeksion, TipiInjeksionit, CmimInjeksion, Gjendje from Injeksionet,TipInjeksioni where TipInjeksioni.IdTipInjeksioni = Injeksionet.IdTipi "
    Else
        Adodc5.RecordSource = "select IdInjeksion, EmerInjeksion, TipiInjeksionit, CmimInjeksion, Gjendje from Injeksionet,TipInjeksioni where TipInjeksioni.IdTipInjeksioni = Injeksionet.IdTipi And EmerInjeksion = '" & Combo14 & "'"
    End If
    
    Combo14.Refresh
    Adodc5.Refresh
DataGrid5.Refresh
End Sub

'Gjenerimi i raportit

Private Sub Command47_Click()
InjeksioneRaport.Show
End Sub

Private Sub Command43_Click()
Unload Me
End Sub

     '======================================================== FUND INJEKSIONET ==================================================================================

'LOGOUT
Private Sub Command48_Click()
Unload Me
Login.Show
Login.Text1.Text = ""
Login.Text2.Text = ""
Login.Text1.SetFocus
End Sub

Private Sub Form_Load()
Combo1_Add
Combo2_Add
Combo3_Add
Combo4_Add
Combo6_Add
Combo7_Add
Combo8_Add
Combo10_Add
Combo11_Add
Combo13_Add
Combo14_Add
Combo15_Add
Combo17_Add
DTPicker1.Value = Format(Now(), "dd/MM/yyyy")
DTPicker2.Value = Format(Now(), "dd/MM/yyyy")
DTPicker3.Value = Format(Now(), "dd/MM/yyyy")
DTPicker4.Value = Format(Now(), "dd/MM/yyyy")
End Sub

