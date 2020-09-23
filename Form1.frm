VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "http://www.c2i.fr"
   ClientHeight    =   1215
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1215
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Modify Microsoft Internet Explorer URL"
      Height          =   435
      Left            =   60
      TabIndex        =   0
      Top             =   720
      Width           =   4575
   End
   Begin VB.Label Label1 
      Caption         =   "Current URL :"
      Height          =   195
      Left            =   60
      TabIndex        =   2
      Top             =   60
      Width           =   2535
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   60
      TabIndex        =   1
      Top             =   360
      Width           =   4575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'//****************************************//
'//  Copyright c2i - Richard CLARK
'//  http://www.c2i.fr
'//  rc@c2i.fr
'//  Code élaboré avec c2iExplorer
'//**************************************//

Private Sub Command1_Click()
SetURL "http://www.c2i.fr"
End Sub

Private Sub Form_Load()
lblInfo = GetURL
End Sub
