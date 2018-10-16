VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ProgressForm 
   Caption         =   "Macro Progress..."
   ClientHeight    =   5299
   ClientLeft      =   91
   ClientTop       =   406
   ClientWidth     =   5285
   OleObjectBlob   =   "ProgressForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ProgressForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btn_ok_Click()
    ' Description: Hides the form when the macro is complete
    ' Dependencies: None
    ' Inputs: None
    ' Outputs: None
    On Error GoTo ErrorHandle
    Me.Hide
    Exit Sub
ErrorHandle:
    custErr.RaiseError "ProgressForm - OK Button"
End Sub
