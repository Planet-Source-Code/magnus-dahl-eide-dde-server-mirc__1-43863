VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "mIRC Sender by sPiKie"
   ClientHeight    =   1950
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3705
   LinkTopic       =   "Form1"
   ScaleHeight     =   1950
   ScaleWidth      =   3705
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox mirc 
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   960
      Width           =   3015
   End
   Begin VB.CommandButton SendDDE 
      Caption         =   "Send"
      Height          =   375
      Left            =   1200
      TabIndex        =   0
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Enter the command here: (i.ex. /msg #channel Hey, im using mIRC sender by spikie"
      Height          =   735
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   3255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

End Sub

Private Sub SendDDE_Click()
mirc.LinkTopic = "mirc|command" ' this sets the DDE server that we want to connect to.
mirc.LinkMode = vbLinkNone ' This resets the DDE client.
On Error GoTo NoServer ' Incase there is no DDE server running.
mirc.LinkItem = mirc.Text ' This sets the data that we want to send, to the contents on the textbox.
mirc.LinkMode = vbLinkManual ' This puts the textbox in DDE manual mode.
mirc.LinkPoke ' This sends the command to mirc.
mirc.LinkMode = vbLinkNone ' And then this resets the client again.

Exit Sub

NoServer:
MsgBox "Server not found." 'No server found
End Sub
