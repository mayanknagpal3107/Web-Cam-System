VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add new camera"
   ClientHeight    =   5130
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   4275
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   5130
   ScaleWidth      =   4275
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3300
      TabIndex        =   2
      Top             =   4680
      Width           =   855
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   375
      Left            =   2220
      TabIndex        =   1
      Top             =   4680
      Width           =   855
   End
   Begin VB.ListBox lstFilters 
      Height          =   4470
      Left            =   0
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   0
      Width           =   4275
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Oked As Boolean
Public CameraName As String

Private Sub cmdCancel_Click()
    Oked = False
    lstFilters.SetFocus
    Hide
End Sub

Private Sub cmdOk_Click()
    Oked = True
    With lstFilters
        CameraName = .List(.ListIndex)
    End With
    lstFilters.SetFocus
    Hide
End Sub

Private Sub Form_Activate()
    Dim rfiEach As QuartzTypeLib.IRegFilterInfo
    
    lstFilters.Clear
    With New QuartzTypeLib.FilgraphManager
        For Each rfiEach In .RegFilterCollection
            lstFilters.AddItem rfiEach.Name
        Next
    End With
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = vbFormControlMenu Then
        Cancel = True
        Oked = False
        Hide
    End If
End Sub

Private Sub Form_Resize()
    If WindowState <> vbMinimized Then
        lstFilters.Width = ScaleWidth
    End If
End Sub

Private Sub lstFilters_Click()
    cmdOk.Enabled = lstFilters.ListIndex > -1
End Sub
