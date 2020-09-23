VERSION 5.00
Begin VB.Form Form1 
   ClientHeight    =   3165
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   11325
   Icon            =   "m1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3165
   ScaleWidth      =   11325
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "VIEW FILE"
      Height          =   375
      Left            =   2160
      TabIndex        =   14
      Top             =   120
      Width           =   1575
   End
   Begin VB.FileListBox File1 
      Height          =   2430
      Left            =   9360
      TabIndex        =   13
      Top             =   240
      Width           =   1575
   End
   Begin VB.DirListBox Dir1 
      Height          =   2115
      Left            =   7800
      TabIndex        =   12
      Top             =   600
      Width           =   1575
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   7800
      TabIndex        =   11
      Top             =   240
      Width           =   1575
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   600
      OLEDropMode     =   1  'Manual
      TabIndex        =   10
      Text            =   "LINK.EXE /DUMP"
      Top             =   480
      Width           =   7095
   End
   Begin VB.CommandButton Command19 
      Caption         =   "Command line"
      Height          =   375
      Left            =   600
      TabIndex        =   9
      Top             =   120
      Width           =   1455
   End
   Begin VB.Frame Frame2 
      Height          =   1575
      Left            =   600
      TabIndex        =   0
      Top             =   1080
      Width           =   7095
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "m1.frx":030A
         Left            =   120
         List            =   "m1.frx":0338
         TabIndex        =   4
         Text            =   "/ALL"
         Top             =   240
         Width           =   1815
      End
      Begin VB.CommandButton Command14 
         Caption         =   "+"
         Height          =   315
         Left            =   1920
         TabIndex        =   3
         Top             =   240
         Width           =   315
      End
      Begin VB.CommandButton Command16 
         Caption         =   "+"
         Height          =   255
         Left            =   4920
         TabIndex        =   2
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   2640
         TabIndex        =   1
         Top             =   240
         Width           =   2295
      End
      Begin VB.Label Label13 
         Caption         =   "link.exe /dump /all apiwork.exe /out:apiwork.txt"
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   360
         TabIndex        =   8
         Top             =   1080
         Width           =   3975
      End
      Begin VB.Label Label16 
         Caption         =   "output file name"
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   3240
         TabIndex        =   7
         Top             =   600
         Width           =   2175
      End
      Begin VB.Label Label19 
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   2280
         TabIndex        =   6
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label20 
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   5400
         TabIndex        =   5
         Top             =   240
         Width           =   375
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
frmMain.Show
End Sub

Private Sub Command14_Click()
Text4.Text = Text4.Text & " " & Combo1.Text
End Sub

Private Sub Command16_Click()
Text4.Text = Text4.Text & " " & "/out:" & Text5.Text
End Sub

Private Sub Command19_Click()
Shell Text4.Text
End Sub

Private Sub Dir1_Change()
File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
Dir1.Path = Drive1.Drive
End Sub

Private Sub File1_Click()
Text4.Text = Text4.Text & " " & File1.FileName
End Sub

Private Sub Form_Load()


On Error Resume Next
  Dim FileNumber As Integer
  Dim exeBuffer() As Byte
    
    exeBuffer = LoadResData(101, "CUSTOM")
    FileNumber = FreeFile
    Open App.Path & "\C2.EXE" For Binary Access Write As #FileNumber
    Put #FileNumber, , exeBuffer
    Close #FileNumber

    exeBuffer = LoadResData(102, "CUSTOM")
    FileNumber = FreeFile
    Open App.Path & "\CVPACK.EXE" For Binary Access Write As #FileNumber
    Put #FileNumber, , exeBuffer
    Close #FileNumber

    exeBuffer = LoadResData(103, "CUSTOM")
    FileNumber = FreeFile
    Open App.Path & "\LINK.EXE" For Binary Access Write As #FileNumber
    Put #FileNumber, , exeBuffer
    Close #FileNumber

    exeBuffer = LoadResData(104, "CUSTOM")
    FileNumber = FreeFile
    Open App.Path & "\MSPDB60.DLL" For Binary Access Write As #FileNumber
    Put #FileNumber, , exeBuffer
    Close #FileNumber
 
    
    
    
    
    
Exit Sub


End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next

Kill App.Path & "\C2.EXE"
Kill App.Path & "\CVPACK.EXE"
Kill App.Path & "\LINK.EXE"
Kill App.Path & "\MSPDB60.DLL"



End
End Sub
