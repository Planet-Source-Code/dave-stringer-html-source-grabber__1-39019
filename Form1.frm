VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form1 
   Caption         =   "HTML Source Grabber"
   ClientHeight    =   5295
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   8655
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5295
   ScaleWidth      =   8655
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog cdialog 
      Left            =   4080
      Top             =   2400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   480
      TabIndex        =   0
      Text            =   "c6d.net"
      Top             =   0
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Get It"
      Default         =   -1  'True
      Height          =   255
      Left            =   7560
      TabIndex        =   2
      Top             =   20
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   2880
      TabIndex        =   1
      Text            =   "http://www.c6d.net/index.html"
      Top             =   0
      Width           =   4575
   End
   Begin RichTextLib.RichTextBox Text1 
      Height          =   4935
      Left            =   0
      TabIndex        =   3
      Top             =   360
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   8705
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   2
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"Form1.frx":030A
   End
   Begin MSWinsockLib.Winsock sckTCP 
      Left            =   3960
      Top             =   3600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label Label2 
      Caption         =   "Host:"
      Height          =   255
      Left            =   80
      TabIndex        =   5
      Top             =   50
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "URL:"
      Height          =   255
      Left            =   2480
      TabIndex        =   4
      Top             =   45
      Width           =   375
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuSAVE 
         Caption         =   "Save As"
      End
      Begin VB.Menu seperator 
         Caption         =   "-"
      End
      Begin VB.Menu mnuexit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Text1.SetFocus
Text1.Text = ""
sckTCP.Close
sckTCP.RemoteHost = Text3
sckTCP.RemotePort = 80
sckTCP.Connect
Do
If sckTCP.State = 7 Then Exit Do
DoEvents
Loop
sckTCP.SendData "GET " & Text2 & " HTTP/1.0" & vbCrLf & "Accept: */*" & vbCrLf & "Accept: text/html" & vbCrLf & vbCrLf
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
sckTCP.Close
End Sub
Private Sub Form_Resize()
On Error Resume Next
Text1.Width = Me.ScaleWidth
Text1.Height = Me.ScaleHeight - 370
Text1.Top = 360
Text2.Width = Me.Width - Text2.Left - Command1.Width - 300
Command1.Left = Text2.Width + Text3.Width + Label1.Width + Label2.Width + 400
End Sub

Private Sub mnuexit_Click()
Unload Me
End Sub

Private Sub mnuSAVE_Click()
cdialog.ShowSave
fts = cdialog.FileName
Text1.SaveFile fts, 1
End Sub

Private Sub sckTCP_DataArrival(ByVal bytesTotal As Long)
sckTCP.GetData temp$
Text1.SelText = temp$
Text1.SelStart = Len(Text1)
End Sub


