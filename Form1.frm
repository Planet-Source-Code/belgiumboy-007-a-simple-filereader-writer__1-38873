VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5100
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6270
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5100
   ScaleWidth      =   6270
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   3840
      TabIndex        =   4
      Text            =   "Test.hello"
      Top             =   4440
      Width           =   2415
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   0
      TabIndex        =   3
      Text            =   "C:\"
      Top             =   4800
      Width           =   6255
   End
   Begin VB.CommandButton Command2 
      Caption         =   "read file"
      Height          =   375
      Left            =   1200
      TabIndex        =   2
      Top             =   4320
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "write to file"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   4320
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   4215
      Left            =   0
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   0
      Width           =   6255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private fso As New FileSystemObject
Private strName As String

Private Sub Command1_Click()
    With fso
        strName = .BuildPath(Text2.Text, Text3.Text)
        Set strm = .CreateTextFile(strName, True)
        strm.Write (Text1.Text)
    End With
    MsgBox "text saved"
    Exit Sub
a:
MsgBox "an error has occurd"
End Sub

Private Sub Command2_Click()
On Error GoTo a
    With fso
        strName = Text2.Text & "\" & Text3.Text
        Set strm = .OpenTextFile(strName, ForReading)
        With strm
            Do Until .AtEndOfStream
                Text1.Text = Text1.Text & .ReadLine & vbCrLf
            Loop
        End With
    End With
    Exit Sub
a:
MsgBox "an error has occurd"
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub
