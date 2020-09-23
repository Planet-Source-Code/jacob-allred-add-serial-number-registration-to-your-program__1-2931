VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SmokeCrypt"
   ClientHeight    =   4905
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6855
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4905
   ScaleWidth      =   6855
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command4 
      Caption         =   "&Copy Output"
      Height          =   375
      Left            =   3720
      TabIndex        =   7
      Top             =   4440
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Paste to Input"
      Height          =   375
      Left            =   5040
      TabIndex        =   6
      Top             =   4440
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Decrypt"
      Height          =   375
      Left            =   1920
      TabIndex        =   4
      Top             =   4440
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Encrypt"
      Height          =   375
      Left            =   600
      TabIndex        =   2
      Top             =   4440
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   4005
      Left            =   3480
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   360
      Width           =   3255
   End
   Begin VB.TextBox Text1 
      Height          =   4005
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   360
      Width           =   3255
   End
   Begin VB.Label Label2 
      Caption         =   "Output:"
      Height          =   255
      Left            =   3480
      TabIndex        =   5
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Input:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
   Public Function CryptString2(txtString As String, EnCrypt As Boolean) As String


       On Error GoTo errhandler
       Dim x As Integer
       Dim outString As String
       Dim iLen As Integer
       Dim sFirstSeed As String
       Dim sSecondSeed As String
       Dim iSeed As Integer


       If EnCrypt Then
           sFirstSeed = Left(txtString, 1)
           sSecondSeed = Mid(txtString, 2, 1)
           iSeed = (Asc(sFirstSeed) + Asc(sSecondSeed)) Mod 2
           iLen = Len(txtString)


           For x = 1 To iLen
               outString = Chr((Asc(Mid$(txtString, x, 1)) Xor iSeed) + 2) & outString
           Next


           outString = Chr(Asc(sFirstSeed) * 2 + 3) & outString
           outString = outString & Chr(Asc(sSecondSeed) * 2 - 3)
       Else
           sFirstSeed = Chr((Asc(Left(txtString, 1)) - 3) \ 2)
           sSecondSeed = Chr((Asc(Right(txtString, 1)) + 3) \ 2)
           iSeed = (Asc(sFirstSeed) + Asc(sSecondSeed)) Mod 2
           iLen = Len(txtString) - 1


           For x = 2 To iLen
               outString = Chr((Asc(Mid$(txtString, x, 1)) Xor iSeed) - 2) & outString
           Next


       End If


       CryptString2 = outString
       Exit Function
errhandler:
       MsgBox "Error in SmokeCrypt" & vbCrLf & "Error: " & Err.Description & vbCrLf & "Number: " & Err.Number
       CryptString2 = ""
   End Function

Sub command1_click()
     Text2.Text = CryptString2(Text1.Text, True)
   End Sub

Private Sub Command2_Click()
     Text2.Text = CryptString2(Text1.Text, False)
End Sub

Private Sub Command3_Click()
Text1.Text = Clipboard.GetText
End Sub

Private Sub Command4_Click()
Clipboard.Clear ' Clear Clipboard.
Clipboard.SetText Text2.Text    ' Put text on Clipboard.
End Sub
