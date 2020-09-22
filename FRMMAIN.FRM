VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "IP Resolver"
   ClientHeight    =   2280
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5190
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2280
   ScaleWidth      =   5190
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtIP 
      BackColor       =   &H8000000F&
      Height          =   285
      Left            =   120
      Locked          =   -1  'True
      MouseIcon       =   "frmMain.frx":0742
      MousePointer    =   99  'Custom
      TabIndex        =   10
      TabStop         =   0   'False
      Text            =   "Unresolved..."
      Top             =   960
      Width           =   1815
   End
   Begin VB.Timer tmrOut 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   4080
      Top             =   120
   End
   Begin MSWinsockLib.Winsock wskResolve 
      Left            =   4560
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton cmdResolve 
      Caption         =   "&Resolve"
      Height          =   320
      Left            =   2040
      TabIndex        =   2
      Top             =   960
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "Common Ports"
      Height          =   1455
      Left            =   3240
      TabIndex        =   9
      Top             =   720
      Width           =   1815
      Begin VB.CommandButton cmdPOP3 
         Caption         =   "POP3"
         Height          =   375
         Left            =   960
         TabIndex        =   6
         Top             =   840
         Width           =   735
      End
      Begin VB.CommandButton cmdSMTP 
         Caption         =   "SMTP"
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   840
         Width           =   735
      End
      Begin VB.CommandButton cmdFTP 
         Caption         =   "FTP"
         Height          =   375
         Left            =   960
         TabIndex        =   4
         Top             =   360
         Width           =   735
      End
      Begin VB.CommandButton cmdHTTP 
         Caption         =   "HTTP"
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.TextBox txtPort 
      Height          =   285
      Left            =   3240
      TabIndex        =   1
      Text            =   "80"
      Top             =   360
      Width           =   1815
   End
   Begin VB.TextBox txtHostName 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Text            =   "www.planet-source-code.com"
      Top             =   360
      Width           =   2895
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Simple Host Resolver Example By: Chuck Hall"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   240
      TabIndex        =   12
      Top             =   1440
      Width           =   2655
   End
   Begin VB.Label Label3 
      Caption         =   "IP Address:"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   720
      Width           =   2895
   End
   Begin VB.Label Label2 
      Caption         =   "Port:"
      Height          =   255
      Left            =   3240
      TabIndex        =   8
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Enter a host name to resolve:"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   2895
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00008000&
      FillStyle       =   0  'Solid
      Height          =   735
      Left            =   120
      Top             =   1440
      Width           =   2895
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Use this however you want to :c)

Private Sub cmdFTP_Click()

    txtPort.Text = 21

End Sub

Private Sub cmdHTTP_Click()

    txtPort.Text = 80

End Sub

Private Sub cmdPOP3_Click()

    txtPort.Text = 110

End Sub

Private Sub cmdResolve_Click()

    'This is the sub that gets the resolving process started...
    'We have to make the winsock control establish a connection
    'with the computer we want to resolve...

    txtIP.Text = "Resolving..."
    
    'First disable this command button so they
    'don't click again until we get this test done
    cmdResolve.Enabled = False
    
    'Now connect the winsock
    wskResolve.RemoteHost = txtHostName.Text
    wskResolve.RemotePort = txtPort.Text
    wskResolve.Connect
    
    'Now enable our timer incase a connection isn't made
    'our timer will return an error instead of letting us
    'sit here forever wondering...
    tmrOut.Enabled = True

End Sub

Private Sub cmdSMTP_Click()

    txtPort.Text = 25

End Sub

Private Sub tmrOut_Timer()
    
    'This will only happen if we can't connect to the host
    'So we need to return an error
    txtIP.Text = "Error Connecting..."
    
    'Then we'll close up our winsock control so it stops
    'Whatever it's doing
    wskResolve.Close
    
    'Disable this timer because we don't need it turned on now
    tmrOut.Enabled = False
    
    'Re-enable the command button so they can try again
    cmdResolve.Enabled = True
    
End Sub

Private Sub txtHostName_GotFocus()

    'This checks if the default value is in there
    'if so highlight it for easy overwriting
    If txtHostName.Text = "www.planet-source-code.com" Then
        txtHostName.SelStart = 0
        txtHostName.SelLength = Len(txtHostName.Text)
    End If

End Sub

Private Sub txtIP_Click()

    'Copy the resolved IP to the clipboard if they click
    Clipboard.SetText txtIP.Text

End Sub

Private Sub wskResolve_Connect()

    'This gets triggered when we successfully connect to the host
    'At this point, the winsock control has automatically resolved
    'the IP to the RemoteHostIP property so we just get that
    
    'Disable our time out timer since we connected
    tmrOut.Enabled = False
    
    'Set our IP Box to say the new IP
    txtIP.Text = wskResolve.RemoteHostIP
    
    'Now close the winsock so it's ready for next time
    wskResolve.Close
    
    'Re-enable our command button in case they want to try again
    cmdResolve.Enabled = True

End Sub
