VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4500
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5265
   LinkTopic       =   "Form1"
   ScaleHeight     =   4500
   ScaleWidth      =   5265
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox txtUri 
      Height          =   315
      ItemData        =   "Form1.frx":0000
      Left            =   120
      List            =   "Form1.frx":0010
      TabIndex        =   18
      Text            =   "ftp://myName:myPass@host.name.com/resource.blah"
      Top             =   480
      Width           =   5055
   End
   Begin VB.TextBox txtLongResource 
      Height          =   285
      Left            =   1320
      TabIndex        =   17
      Top             =   3960
      Width           =   3855
   End
   Begin VB.TextBox txtQuery 
      Height          =   285
      Left            =   1320
      TabIndex        =   16
      Top             =   3600
      Width           =   3855
   End
   Begin VB.TextBox txtResource 
      Height          =   285
      Left            =   1320
      TabIndex        =   15
      Top             =   3240
      Width           =   3855
   End
   Begin VB.TextBox txtPassword 
      Height          =   285
      Left            =   1320
      TabIndex        =   14
      Top             =   2880
      Width           =   3855
   End
   Begin VB.TextBox txtUsername 
      Height          =   285
      Left            =   1320
      TabIndex        =   13
      Top             =   2520
      Width           =   3855
   End
   Begin VB.TextBox txtPort 
      Height          =   285
      Left            =   1320
      TabIndex        =   12
      Top             =   2160
      Width           =   3855
   End
   Begin VB.TextBox txtHost 
      Height          =   285
      Left            =   1320
      TabIndex        =   11
      Top             =   1800
      Width           =   3855
   End
   Begin VB.TextBox txtProtocol 
      Height          =   285
      Left            =   1320
      TabIndex        =   10
      Top             =   1440
      Width           =   3855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Parse URI"
      Height          =   375
      Left            =   1320
      TabIndex        =   0
      Top             =   840
      Width           =   2415
   End
   Begin VB.Label Label9 
      Caption         =   "Password"
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   2880
      Width           =   2055
   End
   Begin VB.Label Label8 
      Caption         =   "Username"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   2520
      Width           =   1935
   End
   Begin VB.Label Label7 
      Caption         =   "FullResource"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   3960
      Width           =   1815
   End
   Begin VB.Label Label6 
      Caption         =   "QueryString"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   3600
      Width           =   1695
   End
   Begin VB.Label Label5 
      Caption         =   "Resource"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   3240
      Width           =   1575
   End
   Begin VB.Label Label4 
      Caption         =   "Port"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   2160
      Width           =   1575
   End
   Begin VB.Label Label3 
      Caption         =   "Host"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Protocol"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   5160
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Label Label1 
      Caption         =   "enter URI:"
      Height          =   255
      Left            =   480
      TabIndex        =   1
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim uri As New CUri
    uri = txtUri.Text
    txtProtocol = uri.Protocol
    txtHost = uri.Host
    txtPort = uri.Port
    txtUsername = uri.Username
    txtPassword = uri.Password
    txtResource = uri.Resource
    txtQuery = uri.QueryString
    txtLongResource = uri.FullResource
End Sub
