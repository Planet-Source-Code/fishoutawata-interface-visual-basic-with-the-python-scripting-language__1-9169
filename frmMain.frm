VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Python Test"
   ClientHeight    =   1740
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1740
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "&Test Call"
      Height          =   525
      Left            =   3390
      TabIndex        =   0
      Top             =   1140
      Width           =   1245
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

Dim objPythonVB As Object

Set objPythonVB = CreateObject("PythonVB.Demo")
retval = objPythonVB.SayHello()

MsgBox retval

End Sub
