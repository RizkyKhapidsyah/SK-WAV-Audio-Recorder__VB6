VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "WAV Recorder"
   ClientHeight    =   2745
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3870
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2745
   ScaleWidth      =   3870
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command5 
      Caption         =   "Statistics"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Save As..."
      Enabled         =   0   'False
      Height          =   255
      Left            =   2520
      TabIndex        =   3
      Top             =   720
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   2640
      Top             =   2400
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3120
      Top             =   2400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Play"
      Enabled         =   0   'False
      Height          =   495
      Left            =   2520
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Stop"
      Enabled         =   0   'False
      Height          =   495
      Left            =   1320
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Record"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Height          =   1335
      Left            =   120
      TabIndex        =   5
      Top             =   1320
      Width           =   3615
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   2295
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()


CommonDialog1.CancelError = True
On Error GoTo ErrHandler1
    CommonDialog1.Filter = "WAV file (*.wav*)|*.wav"
    CommonDialog1.Flags = &H2 Or &H400
    CommonDialog1.ShowSave


'If file already exists then remove it
 FileFound CommonDialog1.FileName
 If ValidFile = True Then
  Kill CommonDialog1.FileName
 End If

'MCI command to save the WAV file
     i = mciSendString("save capture " & CommonDialog1.FileName, 0&, 0, 0)

ErrHandler1:
End Sub

Private Sub Command2_Click()

'Samples Per Second that are supported:
'11025       low quality
'22050       medium quality
'44100     high quality (CD music quality)
 
 
'Bits per sample is 16 or 8


'Channels are 1 (mono) or 2 (stereo)
 
 i = mciSendString("seek capture to start", 0&, 0, 0) 'Always start at the beginning
 i = mciSendString("set capture samplespersec 44100", 0&, 0, 0) 'CD Quality
 i = mciSendString("set capture bitspersample 16", 0&, 0, 0)  '16 bits for better sound
 i = mciSendString("set capture channels 2", 0&, 0, 0) ' 2 channels for stereo
 i = mciSendString("record capture", 0&, 0, 0)  'Start the recording

Command3.Enabled = True  'Enable the STOP BUTTON
Command4.Enabled = False  'Disable the "PLAY" button
Command1.Enabled = False  'Disable the "SAVE AS" button
End Sub

Private Sub Command3_Click()
  i = mciSendString("stop capture", 0&, 0, 0)

Command1.Enabled = True 'Enable the "SAVE AS" button
Command4.Enabled = True 'Enable the "PLAY" button


End Sub


Private Sub Command4_Click()
  i = mciSendString("play capture from 0", 0&, 0, 0)
End Sub


Private Sub Command5_Click()
Dim msg As String
Dim mssg As String * 255

  i = mciSendString("set capture time format ms", 0&, 0, 0)
  i = mciSendString("status capture length", mssg, 255, 0)
msg = "Milliseconds = " & Str(mssg) & vbCrLf

  i = mciSendString("set capture time format bytes", 0&, 0, 0)
  i = mciSendString("status capture length", mssg, 255, 0)
msg = msg & "Bytes = " & Str(mssg) & vbCrLf


i = mciSendString("status capture channels", mssg, 255, 0)
If Str(mssg) = 1 Then
   msg = msg & "Channels = 1 (mono)" & vbCrLf
ElseIf Str(mssg) = 2 Then
   msg = msg & "Channels = 2 (stereo)" & vbCrLf
End If

i = mciSendString("status capture bitspersample", mssg, 255, 0)
   msg = msg & "Bits per sample = " & Str(mssg) & vbCrLf

i = mciSendString("status capture bytespersec", mssg, 255, 0)
   msg = msg & "Bytes per second = " & Str(mssg) & vbCrLf


Label3.Caption = msg

End Sub

Private Sub Form_Load()

 'Close any MCI operations from previous VB programs
 i = mciSendString("close all", 0&, 0, 0)
 
 'Open a new WAV with MCI Command...
 i = mciSendString("open new type waveaudio alias capture", 0&, 0, 0)

End Sub

Private Sub Form_Unload(Cancel As Integer)
 i = mciSendString("close capture", 0&, 0, 0)
End Sub


Private Sub Timer1_Timer()
Dim mssg As String * 255

i = mciSendString("status capture mode", mssg, 255, 0)
Label1.Caption = " " & mssg
End Sub


