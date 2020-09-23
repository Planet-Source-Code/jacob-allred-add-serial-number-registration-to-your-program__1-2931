VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Register"
   ClientHeight    =   1455
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3150
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1455
   ScaleWidth      =   3150
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "E&xit"
      Height          =   375
      Left            =   1680
      TabIndex        =   3
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Register"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   960
      Width           =   1215
   End
   Begin VB.TextBox user1 
      Height          =   285
      Left            =   1200
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
   Begin VB.TextBox pass1 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1200
      TabIndex        =   1
      Top             =   480
      Width           =   1815
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Serial Number:"
      Height          =   255
      Left            =   -120
      TabIndex        =   5
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Name:"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   120
      Width           =   855
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
      
       End If


       CryptString2 = outString
       Exit Function
errhandler:
       MsgBox "Error in SmokeCrypt" & vbCrLf & "Error: " & Err.Description & vbCrLf & "Number: " & Err.Number
       CryptString2 = ""
   End Function

Private Sub Command1_Click()
adduser = CryptString2(user1.Text, True)
If adduser = pass1.Text Then
MsgBox "Thank You For Registering!"
Else
MsgBox "Invalid registration info."
End If
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

