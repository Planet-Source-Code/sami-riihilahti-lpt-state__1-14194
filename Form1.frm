VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "LPT-STATE"
   ClientHeight    =   3090
   ClientLeft      =   3645
   ClientTop       =   2325
   ClientWidth     =   3120
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   3120
   Begin VB.Frame Frame1 
      Caption         =   "Port"
      Height          =   1095
      Left            =   0
      TabIndex        =   0
      Top             =   1980
      Width           =   1695
      Begin VB.OptionButton Option2 
         Caption         =   "LPT2 (278H)"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   600
         Width           =   1275
      End
      Begin VB.OptionButton Option1 
         Caption         =   "LPT1 (378H)"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   300
         Value           =   -1  'True
         Width           =   1275
      End
   End
   Begin MSComctlLib.Slider Slider1 
      Height          =   255
      Left            =   1800
      TabIndex        =   8
      Top             =   2400
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   450
      _Version        =   393216
      SmallChange     =   5
      Min             =   1
      Max             =   100
      SelStart        =   1
      TickFrequency   =   5
      Value           =   1
   End
   Begin VB.Label Label12 
      Caption         =   "Busy"
      Height          =   195
      Left            =   420
      TabIndex        =   15
      Top             =   1560
      Width           =   735
   End
   Begin VB.Image LPTbusyled 
      Height          =   195
      Left            =   120
      Picture         =   "Form1.frx":030A
      Top             =   1560
      Width           =   195
   End
   Begin VB.Label Label11 
      Caption         =   "Ack"
      Height          =   195
      Left            =   1560
      TabIndex        =   14
      Top             =   1320
      Width           =   735
   End
   Begin VB.Image LPTackled 
      Height          =   195
      Left            =   1260
      Picture         =   "Form1.frx":06A1
      Top             =   1320
      Width           =   195
   End
   Begin VB.Label Label10 
      Caption         =   "Paper"
      Height          =   195
      Left            =   420
      TabIndex        =   13
      Top             =   1320
      Width           =   735
   End
   Begin VB.Image LPTpaperled 
      Height          =   195
      Left            =   120
      Picture         =   "Form1.frx":0A38
      Top             =   1320
      Width           =   195
   End
   Begin VB.Label Label9 
      Caption         =   "Select"
      Height          =   195
      Left            =   1560
      TabIndex        =   12
      Top             =   1080
      Width           =   735
   End
   Begin VB.Image LPTselectled 
      Height          =   195
      Left            =   1260
      Picture         =   "Form1.frx":0DCF
      Top             =   1080
      Width           =   195
   End
   Begin VB.Label Label8 
      Caption         =   "Error"
      Height          =   195
      Left            =   420
      TabIndex        =   11
      Top             =   1080
      Width           =   735
   End
   Begin VB.Image LPTerrorled 
      Height          =   195
      Left            =   120
      Picture         =   "Form1.frx":1166
      Top             =   1080
      Width           =   195
   End
   Begin VB.Image led_yellow 
      Height          =   195
      Left            =   2760
      Picture         =   "Form1.frx":14FD
      Top             =   540
      Width           =   195
      Visible         =   0   'False
   End
   Begin VB.Label Label7 
      Caption         =   "msec"
      Height          =   255
      Left            =   1860
      TabIndex        =   10
      Top             =   2760
      Width           =   435
   End
   Begin VB.Label Label6 
      Caption         =   "Update interval:"
      Height          =   255
      Left            =   1860
      TabIndex        =   9
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "Select in"
      Height          =   195
      Left            =   420
      TabIndex        =   6
      Top             =   720
      Width           =   735
   End
   Begin VB.Image LPTselectinled 
      Height          =   195
      Left            =   120
      Picture         =   "Form1.frx":188D
      Top             =   720
      Width           =   195
   End
   Begin VB.Image LPTinitled 
      Height          =   195
      Left            =   120
      Picture         =   "Form1.frx":1C24
      Top             =   480
      Width           =   195
   End
   Begin VB.Label Label4 
      Caption         =   "Init"
      Height          =   195
      Left            =   420
      TabIndex        =   5
      Top             =   480
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "AutoFeed XT"
      Height          =   195
      Left            =   1560
      TabIndex        =   4
      Top             =   720
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Strobe"
      Height          =   195
      Left            =   1560
      TabIndex        =   3
      Top             =   480
      Width           =   615
   End
   Begin VB.Image LPTautofeedled 
      Height          =   195
      Left            =   1260
      Picture         =   "Form1.frx":1FBB
      Top             =   720
      Width           =   195
   End
   Begin VB.Image LPTstrobeled 
      Height          =   195
      Left            =   1260
      Picture         =   "Form1.frx":2352
      Top             =   480
      Width           =   195
   End
   Begin VB.Label Label1 
      Caption         =   "MSB"
      Height          =   195
      Left            =   2100
      TabIndex        =   2
      Top             =   120
      Width           =   435
   End
   Begin VB.Image LPTdataled 
      Height          =   195
      Index           =   7
      Left            =   1800
      Picture         =   "Form1.frx":26E9
      Top             =   120
      Width           =   195
   End
   Begin VB.Image LPTdataled 
      Height          =   195
      Index           =   6
      Left            =   1560
      Picture         =   "Form1.frx":2A80
      Top             =   120
      Width           =   195
   End
   Begin VB.Image LPTdataled 
      Height          =   195
      Index           =   5
      Left            =   1320
      Picture         =   "Form1.frx":2E17
      Top             =   120
      Width           =   195
   End
   Begin VB.Image LPTdataled 
      Height          =   195
      Index           =   4
      Left            =   1080
      Picture         =   "Form1.frx":31AE
      Top             =   120
      Width           =   195
   End
   Begin VB.Image LPTdataled 
      Height          =   195
      Index           =   3
      Left            =   840
      Picture         =   "Form1.frx":3545
      Top             =   120
      Width           =   195
   End
   Begin VB.Image LPTdataled 
      Height          =   195
      Index           =   2
      Left            =   600
      Picture         =   "Form1.frx":38DC
      Top             =   120
      Width           =   195
   End
   Begin VB.Image LPTdataled 
      Height          =   195
      Index           =   1
      Left            =   360
      Picture         =   "Form1.frx":3C73
      Top             =   120
      Width           =   195
   End
   Begin VB.Image LPTdataled 
      Height          =   195
      Index           =   0
      Left            =   120
      Picture         =   "Form1.frx":400A
      Top             =   120
      Width           =   195
   End
   Begin VB.Image led_gray 
      Height          =   195
      Left            =   2760
      Picture         =   "Form1.frx":43A1
      Top             =   60
      Width           =   195
      Visible         =   0   'False
   End
   Begin VB.Image led_green 
      Height          =   195
      Left            =   2760
      Picture         =   "Form1.frx":4738
      Top             =   300
      Width           =   195
      Visible         =   0   'False
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Unload(Cancel As Integer)
UserCancel = True 'Set this to exit this application properly
End Sub

Private Sub Option1_Click()
LPTport = 888 '888 is LPT1 port address
End Sub

Private Sub Option2_Click()
LPTport = 632 '632 is LPT2 port address
End Sub

Private Sub Slider1_Change()
SleepValue = Slider1.Value 'Update pause value
End Sub

