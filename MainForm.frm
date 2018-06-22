VERSION 5.00
Begin VB.Form MainForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Application Window Manager"
   ClientHeight    =   3915
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7590
   Icon            =   "MainForm.frx":0000
   LinkTopic       =   "MainForm"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3915
   ScaleWidth      =   7590
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "List of Windows"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   7335
      Begin VB.ListBox List1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2460
         Left            =   120
         TabIndex        =   0
         Top             =   360
         Width           =   7095
      End
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Maximize"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3120
      Width           =   1695
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Minimize"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3120
      Width           =   1575
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Show"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Hide"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3120
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Refresh"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3120
      Width           =   1455
   End
   Begin VB.Label lbloput 
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Visible         =   0   'False
      Width           =   4695
   End
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Command1_Click()
    Dim u&
    u& = FillTaskListBox(List1)
End Sub

Private Sub Command2_Click()
   Dim hWnd&
    Dim u&
    Dim flag&
    
    hWnd& = RetHandle(lbloput)
    'MsgBox (hWnd&)
    ' hide the window using ShowWindow
    u& = ShowWindow(hWnd&, SW_HIDE)
End Sub

Private Sub Command3_Click()
 Dim hWnd&
    Dim u&
    Dim flag&
    
    hWnd& = RetHandle(lbloput)
    'MsgBox (hWnd&)
    ' hide the window using ShowWindow
    u& = ShowWindow(hWnd&, SW_MINIMIZE)
End Sub

Private Sub Command4_Click()
 Dim hWnd&
    Dim u&
    Dim flag&
    
    hWnd& = RetHandle(lbloput)
    'MsgBox (hWnd&)
    ' hide the window using ShowWindow
    u& = ShowWindow(hWnd&, SW_MAXIMIZE)
End Sub

Private Sub Command5_Click()
Dim hWnd&
    Dim u&
    Dim flag&
    
    hWnd& = RetHandle(lbloput)
    'MsgBox (hWnd&)
    ' hide the window using ShowWindow
    u& = ShowWindow(hWnd&, SW_SHOW)
    u& = ShowWindow(hWnd&, SW_RESTORE)
End Sub

Private Sub Form_Load()
    Dim u&
    u& = FillTaskListBox(List1)
End Sub

Private Sub List1_Click()
lbloput = List1.List(List1.ListIndex)
End Sub
