VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6285
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7320
   LinkTopic       =   "Form1"
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   6285
   ScaleWidth      =   7320
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   3720
      TabIndex        =   1
      Text            =   "Hello!"
      Top             =   2880
      Width           =   495
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "X"
      Height          =   250
      Left            =   6600
      TabIndex        =   0
      Top             =   160
      Width           =   250
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ShapeTheForm As clsTransForm 'make a reference to the class

Private Sub cmdExit_Click()
   
Unload Form1

End Sub

Private Sub Form_Load()
Set ShapeTheForm = New clsTransForm 'instantiate the object from the class

ShapeTheForm.ShapeMe Form1, RGB(255, 255, 255) 'do the real work

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

ShapeTheForm.DragForm Me, Button 'move the form

End Sub


Private Sub Form_Unload(Cancel As Integer)

Set ShapeTheForm = Nothing 'destroy the object
Set Form1 = Nothing 'good practice to free resources VB doesn't normally free when you unload a form!

End Sub
