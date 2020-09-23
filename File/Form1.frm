VERSION 5.00
Begin VB.Form Form1 
   ClientHeight    =   5490
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5460
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5490
   ScaleWidth      =   5460
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Retrieve Data"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1425
      TabIndex        =   9
      Top             =   4935
      Width           =   2490
   End
   Begin VB.Frame Frame2 
      Caption         =   "Data To Be Saved"
      Enabled         =   0   'False
      Height          =   1275
      Left            =   150
      TabIndex        =   4
      Top             =   3570
      Width           =   5220
      Begin VB.CommandButton Command1 
         Caption         =   "Store Data !!!"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   420
         TabIndex        =   10
         Top             =   765
         Width           =   4440
      End
      Begin VB.TextBox txtDataValue 
         Height          =   285
         Left            =   3375
         TabIndex        =   8
         Text            =   "TestValue"
         Top             =   315
         Width           =   1680
      End
      Begin VB.TextBox txtDataName 
         Height          =   315
         Left            =   1080
         TabIndex        =   6
         Text            =   "Name1"
         Top             =   315
         Width           =   1245
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Data Value"
         Height          =   195
         Left            =   2475
         TabIndex        =   7
         Top             =   345
         Width           =   795
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Data Name"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   375
         Width           =   810
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "File Path"
      Height          =   3405
      Left            =   135
      TabIndex        =   0
      Top             =   105
      Width           =   5235
      Begin VB.FileListBox File1 
         Height          =   3015
         Left            =   2865
         TabIndex        =   3
         Top             =   300
         Width           =   2265
      End
      Begin VB.DirListBox Dir1 
         Height          =   2565
         Left            =   135
         TabIndex        =   2
         Top             =   735
         Width           =   2700
      End
      Begin VB.DriveListBox Drive1 
         Height          =   315
         Left            =   135
         TabIndex        =   1
         Top             =   315
         Width           =   2715
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim FileName As String
Dim xFreeFile As Integer

If Right(File1.Path, 1) = "\" Then
        FileName = File1.Path & File1.FileName
    Else
        FileName = File1.Path & "\" & File1.FileName
End If

xFreeFile = FreeFile
Open FileName & ":" & txtDataName For Output As xFreeFile
    Write #xFreeFile, txtDataValue.Text
Close xFreeFile
MsgBox "Data Successfully Stored. You can retrive it anytime."
End Sub

Private Sub Command2_Click()
Dim DName As String
Dim FileName As String
Dim Buf
Dim xFreeFile As Integer
If Right(File1.Path, 1) = "\" Then
        FileName = File1.Path & File1.FileName
    Else
        FileName = File1.Path & "\" & File1.FileName
End If

DName = InputBox("Enter Data Name")
xFreeFile = FreeFile
Open FileName & ":" & DName For Input As xFreeFile
    Input #xFreeFile, Buf
Close xFreeFile
If Buf = "" Then
        MsgBox "Can't find DataName : " & DName
    Else
        MsgBox Buf
End If
End Sub

Private Sub Dir1_Change()
    File1.Path = Dir1.Path
    Frame2.Enabled = False
End Sub

Private Sub Drive1_Change()
    Dir1.Path = Drive1.Drive
    Frame2.Enabled = False
End Sub

Private Sub File1_Click()
    Frame2.Enabled = True
End Sub

