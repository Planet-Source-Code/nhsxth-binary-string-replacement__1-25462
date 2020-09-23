VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Binary String Replacer"
   ClientHeight    =   2835
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "frmBSReplace.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2835
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check1 
      Caption         =   "Back up Original File"
      Height          =   255
      Left            =   2040
      TabIndex        =   11
      Top             =   2160
      Width           =   2535
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   120
      TabIndex        =   10
      Text            =   "100"
      Top             =   2520
      Width           =   855
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4200
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Exit"
      Height          =   255
      Left            =   3360
      TabIndex        =   8
      Top             =   2550
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Replace"
      Height          =   255
      Left            =   2040
      TabIndex        =   7
      Top             =   2550
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   120
      TabIndex        =   6
      Top             =   1800
      Width           =   4455
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   4455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Browse"
      Height          =   255
      Left            =   3600
      TabIndex        =   2
      Top             =   720
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   4455
   End
   Begin VB.Label Label4 
      Caption         =   "Buffer Size:"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Replacement Text:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1560
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   "String to Search For:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   "File:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    On Error GoTo errhandle
    
    CommonDialog1.Filter = "All Files (*.*)|*.*"
    CommonDialog1.FileName = ""
    CommonDialog1.ShowOpen
    If CommonDialog1.FileName <> "" Then Text1.Text = CommonDialog1.FileName
    If CommonDialog1.FileName = "" Then Text1.Text = ""
    Exit Sub
errhandle:
    y = MsgBox("Error #" & Err.Number & ", " & Err.Description, vbExclamation, "Error")
    Exit Sub
End Sub
Private Sub Command2_Click()
    On Error GoTo errhandle
    
    Dim FileNum1 As Integer
    Dim FileNum2 As Integer
    Dim fPos As Integer
    Dim i As Integer
    Dim tFileName As String
    Dim Got As String
    Dim tData As String
    
    If Len(Text3.Text) < Len(Text2.Text) Then
        For i = 1 To Len(Text2.Text) - Len(Text3.Text)
            Text3.Text = Text3.Text & " "
        Next
    End If
    FileNum1 = FreeFile
    FileNum2 = FreeFile + 1
    Got = String(1, " ")
    tFileName = Text1.Text & ".TEMP"
    Open Text1.Text For Binary As #FileNum1
    Open tFileName For Binary As #FileNum2
    Do Until EOF(FileNum1)
        tData = ""
        For i = 1 To Val(Text4.Text)
            Get #FileNum1, , Got
            If EOF(FileNum1) Then Exit For
            tData = tData & Got
        Next
        fPos = InStr(tData, Text2.Text)
        If fPos Then Mid(tData, fPos, Len(Text2.Text)) = Text3.Text
        Put #FileNum2, , tData
    Loop
    Close #FileNum1
    Close #FileNum2
    If Check1.Value = 1 Then
        FileCopy Text1.Text, Text1.Text & ".BAK"
        FileCopy tFileName, Text1.Text
    Else
        Kill Text1.Text
        FileCopy tFileName, Text1.Text
    End If
    Kill tFileName
    y = MsgBox("All Matching Cases Replaced", vbInformation, "Done")
    Exit Sub
errhandle:
    y = MsgBox("Error #" & Err.Number & ", " & Err.Description, vbExclamation, "Error")
    Exit Sub
End Sub
Private Sub Command3_Click()
    End
End Sub
Private Sub Text3_KeyDown(KeyCode As Integer, Shift As Integer)
    If Len(Text3.Text) + 1 > Len(Text2.Text) Then Text3.Locked = True
    If KeyCode = vbKeyBack Then Text3.Locked = False
End Sub
