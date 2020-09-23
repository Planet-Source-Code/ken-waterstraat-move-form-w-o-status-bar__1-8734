VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   ClientHeight    =   2250
   ClientLeft      =   120
   ClientTop       =   120
   ClientWidth     =   5250
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   Picture         =   "frmMove.frx":0000
   ScaleHeight     =   2250
   ScaleWidth      =   5250
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "&Exit"
      Height          =   255
      Left            =   3240
      TabIndex        =   2
      Top             =   1080
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&About"
      Height          =   255
      Left            =   2040
      TabIndex        =   1
      Top             =   1080
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Left Click"
      Height          =   255
      Left            =   840
      TabIndex        =   0
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Menu mnuLeftButton 
      Caption         =   "Left Button"
      Visible         =   0   'False
      Begin VB.Menu mnuLeftButton1 
         Caption         =   "You clicked the left button!"
      End
      Begin VB.Menu mnuLeftButton2 
         Caption         =   "Isn't this cool!!"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Dim MoveScreen As Boolean
 Dim CurrX As Integer
 Dim CurrY As Integer
 Dim MousX As Integer
 Dim MousY As Integer
 
Private Sub Command1_Click()
 'popup menu function makes a premade, invisible menu come up, without the use of the actual menu editor
 'first attribute, 0 is the flags
 'second attribute, 840 is the left coordinate you want the menu to appear at
 'third attribute, 1350 is the top coordinate you want the menu to come up at
  PopupMenu Form1.mnuLeftButton, 0, 840, 1350
  
End Sub

Private Sub Command3_Click()
 MsgBox "Move Form Without Status bar" & vbCrLf & vbCrLf & _
   "Programmed by: Kenneth Waterstraat" & vbCrLf & _
   "Programmed in: Visual Basic 6.0" & vbCrLf & _
   "Email: ExoDus3259@aol.com", vbOKOnly + vbInformation, "About Move Form"
   
End Sub

Private Sub Command4_Click()
 Set Form1 = Nothing
 Unload Me
 End
 
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 MoveScreen = True
  MousX = X
  MousY = Y
  
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If MoveScreen Then
  CurrX = Form1.Left - MousX + X
  CurrY = Form1.Top - MousY + Y
   Form1.Move CurrX, CurrY
    End If
    
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
 MoveScreen = False
 
End Sub
