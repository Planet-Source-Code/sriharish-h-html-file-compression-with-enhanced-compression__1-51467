VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4485
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   7380
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmSplash.frx":000C
   ScaleHeight     =   4485
   ScaleWidth      =   7380
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Email Author"
      Height          =   375
      Left            =   1440
      TabIndex        =   1
      Top             =   3600
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "About"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   3600
      Width           =   1095
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub Command1_Click()
MsgBox "HTML Compressor v1.0 with Enhanced Compression" & vbCrLf & vbCrLf & _
        "This program is optimised for HTML files only. It is recommended for you to take backup of your files before compressing." & vbrlf & vbCrLf & "FREEWARE- But people sell it for $29.95, so atleast VOTE FOR ME", vbInformation
               
End Sub

Private Sub Command2_Click()
On Error Resume Next
Dim shellsuccess As Long
shellsuccess = ShellExecute(fH, "Open", "mailto:sriharish@msn.com?Subject=About HTML Compressor", 0&, 0&, 10)
End Sub

Private Sub Form_click()
Form1.Show
Unload Me
End Sub
