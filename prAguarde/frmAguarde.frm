VERSION 5.00
Begin VB.Form frmAguarde 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   1125
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2895
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1125
   ScaleWidth      =   2895
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Command1"
      Height          =   375
      Left            =   1920
      TabIndex        =   2
      Top             =   720
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   720
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   840
      TabIndex        =   0
      Top             =   720
      Width           =   735
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C00000&
      BackStyle       =   1  'Opaque
      Height          =   255
      Left            =   0
      Top             =   240
      Width           =   2775
   End
   Begin VB.Shape Picture1 
      BackColor       =   &H000080FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H000080FF&
      Height          =   255
      Left            =   0
      Top             =   0
      Width           =   2775
   End
End
Attribute VB_Name = "frmAguarde"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim myThread As clsThreads


    
Private Sub Command1_Click()

    
    Thread_Start_PgBar myThread

    
    
End Sub


Private Sub Command2_Click()
     Thread_Stop myThread
    
    'Set myThread = Nothing
End Sub
