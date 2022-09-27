VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form FReplace 
   Caption         =   "Replace"
   ClientHeight    =   6000
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7875
   LinkTopic       =   "Form1"
   Picture         =   "FReplace.frx":0000
   ScaleHeight     =   6000
   ScaleWidth      =   7875
   StartUpPosition =   3  'Windows Default
   Begin RichTextLib.RichTextBox RTF 
      Height          =   495
      Left            =   720
      TabIndex        =   7
      Top             =   4080
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   873
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"FReplace.frx":0342
   End
   Begin VB.FileListBox FileList 
      Height          =   480
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox TFind 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      TabIndex        =   2
      Top             =   2640
      Width           =   4695
   End
   Begin VB.TextBox TReplace 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      TabIndex        =   1
      Top             =   3120
      Width           =   4695
   End
   Begin VB.CommandButton CReplace 
      Caption         =   "Find And Replace"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2640
      TabIndex        =   0
      Top             =   3840
      Width           =   2775
   End
   Begin VB.Label Label3 
      Caption         =   "Give First Common Expression that should be replaced by given Expression."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   960
      TabIndex        =   5
      Top             =   720
      Width           =   6375
   End
   Begin VB.Label Label2 
      Caption         =   "Replace"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   840
      TabIndex        =   4
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Find"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   840
      TabIndex        =   3
      Top             =   2640
      Width           =   1215
   End
End
Attribute VB_Name = "FReplace"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CReplace_Click()

    FileList.Path = App.Path
    FileList.Refresh 'REFRESHES FILE LIST
    i = 0
    While i < FileList.ListCount
        
        If ((InStr(1, FileList.List(i), ".html") > 0) Or (InStr(1, FileList.List(i), ".htm") > 0)) Then
            
            RTF.LoadFile App.Path & "\" & FileList.List(i), vbCFText
            
            While RTF.Find(TFind.Text) > -1
                RTF.SelText = TReplace.Text
            Wend
            RTF.SaveFile (App.Path & "\" & FileList.List(i)), rtfText
            
        End If
        i = i + 1
    Wend
    MsgBox "Replaced All !", vbInformation
End Sub

    
