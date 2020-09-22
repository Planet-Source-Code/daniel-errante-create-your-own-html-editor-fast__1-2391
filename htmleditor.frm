VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "HTML Editor by Daniel"
   ClientHeight    =   5355
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8115
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "htmleditor.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5355
   ScaleWidth      =   8115
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command5 
      Caption         =   "Minimize"
      Height          =   375
      Left            =   2760
      TabIndex        =   5
      Top             =   4920
      Width           =   975
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2640
      Top             =   4920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Open"
      Height          =   375
      Left            =   3840
      TabIndex        =   4
      Top             =   4920
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Save"
      Height          =   375
      Left            =   4920
      TabIndex        =   3
      Top             =   4920
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Preview"
      Height          =   375
      Left            =   6000
      TabIndex        =   2
      Top             =   4920
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "E&xit"
      Height          =   375
      Left            =   7080
      TabIndex        =   1
      Top             =   4920
      Width           =   975
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   4695
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   8281
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"htmleditor.frx":030A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
rep% = MsgBox("Do you want to quit?", vbQuestion + vbYesNo)
If rep% = vbYes Then
End
Else
Exit Sub
End If

End Sub

Private Sub Command2_Click()
Open App.Path & "\preview.html" For Output As #1
Print #1, RichTextBox1.Text
Close #1
Load frmBrowser
frmBrowser.Show
frmBrowser.brwWebBrowser.Navigate App.Path & "\preview.html"

End Sub

Private Sub Command3_Click()
CommonDialog1.Filter = "HTML Files (*.html)|*.html|HTM Files (*.htm)|*.htm)"
CommonDialog1.ShowSave
If CommonDialog1.FileName <> "" Then
    Open CommonDialog1.FileName For Output As #1
    Print #1, RichTextBox1.Text
    Close #1
End If

End Sub

Private Sub Command4_Click()
CommonDialog1.Filter = "HTML Files (*.html)|*.html|HTM Files (*.htm)|*.htm)"
CommonDialog1.ShowOpen
If CommonDialog1.FileName <> "" Then
    Open CommonDialog1.FileName For Input As #1
    Do Until EOF(1)
    Line Input #1, lineoftext$
    alltext$ = alltext$ & lineoftext$
    RichTextBox1.Text = alltext$
    Loop
    Close #1
End If

End Sub

Private Sub Command5_Click()
Me.WindowState = 1

End Sub

Private Sub Form_Load()
RichTextBox1.Text = "<HTML>" & vbCrLf & vbCrLf & "<HEAD>" & vbCrLf & "<TITLE>" & "Web Page</TITLE>" & vbCrLf & "</HEAD>" & vbCrLf & vbCrLf & "<BODY>" & vbCrLf & vbCrLf & "</BODY>" & vbCrLf & vbCrLf & "</HTML>"

End Sub
