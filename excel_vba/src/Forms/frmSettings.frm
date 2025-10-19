VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSettings 
   Caption         =   "Datarails Connection Settings"
   ClientHeight    =   3330
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5580
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   360
      Left            =   2880
      TabIndex        =   4
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Default         =   -1  'True
      Height          =   360
      Left            =   4200
      TabIndex        =   3
      Top             =   2760
      Width           =   1215
   End
   Begin VB.TextBox txtBridgeToken 
      Height          =   360
      Left            =   180
      TabIndex        =   2
      Top             =   2040
      Width           =   5235
   End
   Begin VB.TextBox txtBackendUrl 
      Height          =   360
      Left            =   180
      TabIndex        =   1
      Top             =   960
      Width           =   5235
   End
   Begin VB.Label lblBridgeToken 
      Caption         =   "Bridge admin token (required to store API keys)"
      Height          =   255
      Left            =   180
      TabIndex        =   5
      Top             =   1680
      Width           =   5235
   End
   Begin VB.Label lblBackendUrl 
      Caption         =   "Backend URL (e.g. http://localhost:8000)"
      Height          =   255
      Left            =   180
      TabIndex        =   6
      Top             =   660
      Width           =   5235
   End
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    On Error GoTo HandleError
    Settings.PersistSettings Me.txtBackendUrl.Text, Me.txtBridgeToken.Text
    Unload Me
    Exit Sub
HandleError:
    MsgBox Err.Description, vbCritical + vbOKOnly, "Datarails Settings"
End Sub
