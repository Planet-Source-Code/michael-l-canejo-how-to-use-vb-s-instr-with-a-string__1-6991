VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "How to use Instr"
   ClientHeight    =   870
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   2085
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   870
   ScaleWidth      =   2085
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Text            =   "Mike Canejo"
      Top             =   0
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Get First && Last Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   2055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error GoTo StringError

 TheString$ = Text1
 'To Simplify the term used below

  GetLastName$ = Mid(TheString$, InStr(TheString$, " ") + 1)
  'This will Find the Space(" ") Char and get the text after it
 
   GetFirstName$ = Replace(TheString$, GetLastName$, "")
   'This will find the the string "GetLastName$" in the string "TheString$"
   'And replace it with nothing so it removes itself from TheString$
   'Leaving you with the First Name in Text1
  
     MsgBox "Your first name is: " & GetFirstName$ & vbCrLf & vbCrLf & "Your last name is: " & GetLastName$ _
     , vbSystemModal + vbInformation, "String Finder By: Mike"
     'Displays the First Name and Last Name
   Exit Sub
  
StringError:
 MsgBox Error & vbCrLf & "There was an error searching for the first and last name in the string: " & TheString$ _
 , vbSystemModal + vbCritical, "String Error"
 'If there's an error searching for the first and last name,
 'Go here and display the error
End Sub
