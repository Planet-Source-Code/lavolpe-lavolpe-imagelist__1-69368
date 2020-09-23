VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   3345
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   3345
   StartUpPosition =   3  'Windows Default
   Begin Project1.UserControl1 UserControl13 
      Height          =   1155
      Left            =   105
      TabIndex        =   0
      Top             =   1950
      Width           =   1380
      _ExtentX        =   2434
      _ExtentY        =   2037
      ilName          =   "UserControl11(1)"
      imgID           =   "2"
   End
   Begin Project1.UserControl1 UserControl12 
      Height          =   1230
      Left            =   165
      TabIndex        =   1
      Top             =   600
      Width           =   1350
      _ExtentX        =   2381
      _ExtentY        =   2170
      ilName          =   "UserControl11(0)"
      imgID           =   "6"
   End
   Begin Project1.UserControl1 UserControl11 
      Height          =   585
      Index           =   0
      Left            =   1620
      TabIndex        =   2
      Top             =   120
      Width           =   945
      _ExtentX        =   1667
      _ExtentY        =   1032
      Mode            =   -1
      ilDat1          =   "Form1new.frx":0000
      ilHdr1          =   5
      ilDat2          =   "Form1new.frx":712C
      ilHdr2          =   2
      ilItem1         =   "Form1new.frx":92A7
      ilMisc          =   "Form1new.frx":93C1
   End
   Begin Project1.UserControl1 UserControl11 
      Height          =   585
      Index           =   1
      Left            =   1905
      TabIndex        =   4
      Top             =   750
      Width           =   945
      _ExtentX        =   1667
      _ExtentY        =   1032
      Mode            =   -1
      ilDat1          =   "Form1new.frx":93E9
      ilHdr1          =   4
      ilDat2          =   "Form1new.frx":10D11
      ilHdr2          =   2
      ilItem1         =   "Form1new.frx":10FDA
      ilMisc          =   "Form1new.frx":110D8
   End
   Begin VB.Label Label1 
      Caption         =   "With this example, a custom property page can be used to change images also.  Right click and select properties"
      Height          =   1560
      Left            =   1665
      TabIndex        =   3
      Top             =   1485
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    
End Sub

Private Sub UserControl13_GotFocus()
    
End Sub
