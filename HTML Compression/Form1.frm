VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "HTML Compressor"
   ClientHeight    =   6360
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9660
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6360
   ScaleWidth      =   9660
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check1 
      Caption         =   "Use Enchanced Compression (NOT SAFE)"
      Height          =   255
      Left            =   5880
      TabIndex        =   7
      Top             =   5640
      Width           =   3495
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   720
      Top             =   4680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Select Web Pages"
      MaxFileSize     =   1000
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Clear List"
      Height          =   375
      Left            =   1560
      TabIndex        =   4
      Top             =   5640
      Width           =   1575
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   6120
      Width           =   9735
      _ExtentX        =   17171
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3720
      Top             =   4920
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":144A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Begin Compression"
      Height          =   375
      Left            =   3240
      TabIndex        =   2
      Top             =   5640
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add File"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   5640
      Width           =   1335
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   5415
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "Double Click the list to open HTML file"
      Top             =   120
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   9551
      View            =   3
      Arrange         =   1
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "File"
         Object.Width           =   10584
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Size"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Cpz Ratio %"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Cpz Size"
         Object.Width           =   1499
      EndProperty
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "000"
      Height          =   255
      Left            =   5400
      TabIndex        =   6
      Top             =   5640
      Width           =   735
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Total:"
      Height          =   255
      Left            =   4920
      TabIndex        =   5
      Top             =   5640
      Width           =   495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°
'HTML Compression Code by Sriharish
'
'Email: sriharish@msn.com?Subject=HTML Compressor
'
'Who wants to be a Voter? Its your right to vote, don't lose it
'
'°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Sub Check1_Click()
If MsgBox("Enhanced Compression removes unecessary tags in the HTML file. This is not recommended if you are using Scripts (Like ASP, JAVA, PHP). Only use this compression for only HTML Files. Are you sure you want to continue?", vbExclamation + vbYesNo) = vbYes Then
Check1.Value = 1
Else
Check1.Value = 0
End If
End Sub

Private Sub Command1_Click()
With CommonDialog1
.FileName = ""
.Flags = cdlOFNAllowMultiselect
.Filter = "Html File|*.html|HTM File|*.htm|"
.ShowOpen
filenames = .FileName
End With
If Len(filenames) = 0 Then
Exit Sub
End If
If CommonDialog1.FileTitle <> "" Then
ListView1.ListItems.Add 1, , CommonDialog1.FileName, , 1
ListView1.ListItems(1).SubItems(1) = FileLen(CommonDialog1.FileName)
Label2.Caption = ListView1.ListItems.Count
Exit Sub
End If
spPosition = InStr(filenames, " ")
pathname = Left(filenames, spPosition - 1)
'  Label1.Caption = pathName
filenames = Mid(filenames, spPosition + 1)
' then extract each space delimited file name
If Len(filenames) = 0 Then
List1.AddItem "No files selected"
Label2.Caption = ListView1.ListItems.Count
Exit Sub
Else
spPosition = InStr(filenames, " ")
While spPosition > 0
ListView1.ListItems.Add 1, , pathname & Left(filenames, spPosition - 1), , 1
ListView1.ListItems(1).SubItems(1) = FileLen(pathname & Left(filenames, spPosition - 1))
filenames = Mid(filenames, spPosition + 1)
spPosition = InStr(filenames, " ")
Wend
' Add the last file's name to the list
' (the last file name isn't followed by a space)
ListView1.ListItems.Add 1, , pathname & filenames, , 1
ListView1.ListItems(1).SubItems(1) = FileLen(filenames)
End If
Label2.Caption = ListView1.ListItems.Count
End Sub


Private Sub Command3_Click()
If ListView1.ListItems.Count = 0 Then
MsgBox "No files to Compress. HAW HAW", vbCritical
Exit Sub
End If
If Check1.Value = 1 Then
CompressX
Else
Compressfiles
End If
End Sub
Private Sub Compressfiles()
On Error GoTo error
Dim dat As String
Dim Stock As String
Dim tempdata As String
Dim currentlength As String
Dim Compresslength As String
Dim ratio As String, k As Integer
Dim quote As String
ProgressBar1.Max = ListView1.ListItems.Count
For k = 1 To ListView1.ListItems.Count
ifile = FreeFile
Open ListView1.ListItems(k) For Input As #3
currentlength = LOF(3)
Do Until EOF(3)
Line Input #3, tempdata
If Trim$(tempdata) = "" Then
Else
Stock = Stock & Trim$(tempdata)
End If
Loop
Close #3
Open ListView1.ListItems(k) For Output As #2
Print #2, Stock
Stock = ""
Compresslength = LOF(2)
Close #2
ListView1.ListItems(k).SubItems(3) = Compresslength
ratio = currentlength / Compresslength
ratio = 100 / ratio
ratio = 100 - ratio
ratio = Round(ratio, 2)
ListView1.ListItems(k).SubItems(2) = ratio & " %"
ratio = 0
If ListView1.ListItems.Count = 1 Then
ProgressBar1.Value = ProgressBar1.Max
Else
ProgressBar1.Value = k
End If
Next
Exit Sub
error:
MsgBox Err.Description & vbCrLf & "Source: " & Err.Source & vbCrLf & "Error Code: " & Err.Number, vbCritical
End Sub
Private Sub CompressX()
On Error GoTo error
Dim dat As String
Dim Stock As String
Dim tempdata As String
Dim currentlength As String
Dim Compresslength As String
Dim ratio As String, k As Integer
Dim quote As String
ProgressBar1.Max = ListView1.ListItems.Count
For k = 1 To ListView1.ListItems.Count
ifile = FreeFile
MakeChanges ListView1.ListItems(k), Chr(34), ""
Open ListView1.ListItems(k) For Input As #3
currentlength = LOF(3)
Do Until EOF(3)
Line Input #3, tempdata
If Trim$(tempdata) = "" Then
Else
Stock = Stock & Trim$(tempdata)
End If
Loop
Close #3
Open ListView1.ListItems(k) For Output As #2
Print #2, Stock
Stock = ""
Compresslength = LOF(2)
Close #2
ListView1.ListItems(k).SubItems(3) = Compresslength
ratio = currentlength / Compresslength
ratio = 100 / ratio
ratio = 100 - ratio
ratio = Round(ratio, 2)
ListView1.ListItems(k).SubItems(2) = ratio & " %"
ratio = 0
If ListView1.ListItems.Count = 1 Then
ProgressBar1.Value = ProgressBar1.Max
Else
ProgressBar1.Value = k
End If
Next
Exit Sub
error:
MsgBox Err.Description & vbCrLf & "Source: " & Err.Source & vbCrLf & "Error Code: " & Err.Number, vbCritical
End Sub
Private Sub Command4_Click()
ListView1.ListItems.Clear
End Sub

Private Sub Form_Load()
MsgBox "TIP: Always Backup", vbExclamation
End Sub
Private Sub ListView1_DblClick()
On Error Resume Next
If ListView1.SelectedItem.Text = "" Then
Exit Sub
End If
On Error Resume Next
Dim shellsuccess As Long
shellsuccess = ShellExecute(fH, "Open", ListView1.SelectedItem.Text, 0&, 0&, 10)
End Sub
