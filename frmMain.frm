VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "RTF wrapper class by Adi barda israel"
   ClientHeight    =   5205
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9825
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5205
   ScaleWidth      =   9825
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add new word to editor dictionary"
      Height          =   465
      Left            =   5670
      TabIndex        =   6
      Top             =   4470
      Width           =   2925
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   465
      Left            =   8730
      TabIndex        =   5
      Top             =   4470
      Width           =   1035
   End
   Begin VB.ListBox lstWords 
      Height          =   3765
      Left            =   7260
      TabIndex        =   4
      Top             =   450
      Width           =   2535
   End
   Begin RichTextLib.RichTextBox txtDebug 
      Height          =   930
      Left            =   390
      TabIndex        =   2
      Top             =   3300
      Width           =   6795
      _ExtentX        =   11986
      _ExtentY        =   1640
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"frmMain.frx":0000
   End
   Begin RichTextLib.RichTextBox txtEditor 
      Height          =   2685
      Left            =   360
      TabIndex        =   0
      Top             =   420
      Width           =   6825
      _ExtentX        =   12039
      _ExtentY        =   4736
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"frmMain.frx":009F
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   "Extra words"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   2
      Left            =   7260
      TabIndex        =   7
      Top             =   210
      Width           =   2025
   End
   Begin VB.Label Label1 
      Caption         =   "Debug"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   1
      Left            =   420
      TabIndex        =   3
      Top             =   3090
      Width           =   2025
   End
   Begin VB.Label Label1 
      Caption         =   "Type here VB script code"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   0
      Left            =   390
      TabIndex        =   1
      Top             =   180
      Width           =   2025
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'This is a demo app for my Light RTF Editor class
'I also wrote much more sophisticated class with VB like intelisense !!!
'Enjoy it !!

Option Explicit

Private m_objEditor As CEditor 'main editor object

Private Sub cmdAdd_Click()
    
    Dim s As String
    
    'type the new editor word
    s = Trim$(InputBox("Add new word to editor dictionary"))
    If LenB(s) > 0 Then
        Me.lstWords.AddItem s 'update list box
        m_objEditor.AddEditorWord s & " ", vbBlue 'you can change it to any color you want
        m_objEditor.AddEditorWord s & vbNewLine, vbBlue 'in case you pressed enter
    End If
    
End Sub

Private Sub cmdExit_Click()

    'good bye
    Unload Me
    
End Sub

Private Sub Form_Load()

    Set m_objEditor = New CEditor
    
    'set editor objects
    m_objEditor.SetEditorObjects Me.txtEditor, Me.txtDebug
    
    'hard code init the basic vb script words -
    'you can init any words you want with any colors you like
    m_objEditor.AddEditorWord "Dim ", vbRed
    m_objEditor.AddEditorWord "Select ", vbBlue
    m_objEditor.AddEditorWord "Until ", vbBlue
    m_objEditor.AddEditorWord "Set ", vbBlue
    m_objEditor.AddEditorWord "seq.", vbBlue
    m_objEditor.AddEditorWord "system.", vbBlue
    m_objEditor.AddEditorWord "xml.", vbBlue
    m_objEditor.AddEditorWord "Sub ", vbBlue
    m_objEditor.AddEditorWord "End sub", vbBlue
    m_objEditor.AddEditorWord "For ", vbBlue
    m_objEditor.AddEditorWord "Next ", vbBlue
    m_objEditor.AddEditorWord "Next" & vbNewLine, vbBlue
    m_objEditor.AddEditorWord "Do ", vbBlue
    m_objEditor.AddEditorWord "Do" & vbNewLine, vbBlue
    m_objEditor.AddEditorWord "Loop ", vbBlue
    m_objEditor.AddEditorWord "Loop" & vbNewLine, vbBlue
    m_objEditor.AddEditorWord "If ", vbBlue
    m_objEditor.AddEditorWord "End If ", vbBlue
    m_objEditor.AddEditorWord "End If" & vbNewLine, vbBlue
    m_objEditor.AddEditorWord "End Select ", vbBlue
    m_objEditor.AddEditorWord "End Select" & vbNewLine, vbBlue
    m_objEditor.AddEditorWord "Then ", vbBlue
    m_objEditor.AddEditorWord "Then" & vbNewLine, vbBlue
    m_objEditor.AddEditorWord "Else ", vbBlue
    m_objEditor.AddEditorWord "Else" & vbNewLine, vbBlue

End Sub
