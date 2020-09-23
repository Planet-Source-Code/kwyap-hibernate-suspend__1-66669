VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Hibernate/Suspend"
   ClientHeight    =   1170
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5505
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1170
   ScaleWidth      =   5505
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSuspend 
      Caption         =   "Suspend Now"
      Height          =   495
      Left            =   4080
      TabIndex        =   7
      Top             =   360
      Width           =   1215
   End
   Begin VB.CommandButton cmdHibernate 
      Caption         =   "Hibernate Now"
      Height          =   495
      Left            =   2760
      TabIndex        =   0
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label lbShutdown 
      Caption         =   "Label4"
      Height          =   255
      Left            =   1560
      TabIndex        =   6
      Top             =   840
      Width           =   615
   End
   Begin VB.Label lbSuspend 
      Caption         =   "Label4"
      Height          =   255
      Left            =   1560
      TabIndex        =   5
      Top             =   480
      Width           =   615
   End
   Begin VB.Label lbHibernate 
      Caption         =   "Label4"
      Height          =   255
      Left            =   1560
      TabIndex        =   4
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Can Shutdown :"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Can Suspend :"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Can Hibernate :"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Determines whether the computer supports hibernation
Private Declare Function IsPwrHibernateAllowed Lib "powrprof.dll" () As Long

' Determines whether the computer supports the sleep states
Private Declare Function IsPwrSuspendAllowed Lib "powrprof.dll" () As Long

' Determines whether the computer supports the soft off power state
Private Declare Function IsPwrShutdownAllowed Lib "powrprof.dll" () As Long

' Suspends the system by shutting power down
'   Hibernate - If this parameter is TRUE, the system hibernates. If the parameter is FALSE, the system is suspended
Private Declare Function SetSuspendState Lib "powrprof.dll" (ByVal Hibernate As Long, ByVal ForceCritical As Long, ByVal DisableWakeEvent As Long) As Long

Private Sub cmdHibernate_Click()
    If SetSuspendState(True, True, True) = 0 Then
        MsgBox "Hibernate fails"
    End If
End Sub

Private Sub cmdSuspend_Click()
    If SetSuspendState(False, True, True) = 0 Then
        MsgBox "Suspend fails"
    End If
End Sub

Private Sub Form_Load()
    If IsPwrHibernateAllowed Then
        lbHibernate.Caption = "TRUE"
        cmdHibernate.Enabled = True
    Else
        lbHibernate.Caption = "FALSE"
        cmdHibernate.Enabled = False
    End If
    
    If IsPwrSuspendAllowed Then
        lbSuspend.Caption = "TRUE"
        cmdSuspend.Enabled = True
    Else
        lbSuspend.Caption = "FALSE"
        cmdSuspend.Enabled = False
    End If
    
    If IsPwrShutdownAllowed Then
        lbShutdown.Caption = "TRUE"
    Else
        lbShutdown.Caption = "FALSE"
    End If
End Sub
