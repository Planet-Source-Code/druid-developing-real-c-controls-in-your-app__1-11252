VERSION 5.00
Begin VB.Form frmControls 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Use REAL C++ Controls in your app!!!!!"
   ClientHeight    =   4665
   ClientLeft      =   3780
   ClientTop       =   3210
   ClientWidth     =   7860
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4665
   ScaleWidth      =   7860
End
Attribute VB_Name = "frmControls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'You are wrong here! Go to the Module to see how the code works

Private Sub Form_Unload(Cancel As Integer)
Dim resultMsg As Long
resultMsg = MsgBox("Would you like to vote for my work?", vbInformation Or vbYesNo, "Question")
If resultMsg = vbYes Then Shell "start http://www.planet-source-code.com/vb/scripts/ShowCode.asp?lngWId=1&txtCodeId=11252"
End Sub
