VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmAbout 
   Caption         =   "About iResearch"
   ClientHeight    =   2175
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   6000
   OleObjectBlob   =   "frmAbout.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub UserForm_Activate()
    Dim fso As New FileSystemObject
    Dim ts As TextStream
    Dim version As String
    
    Set ts = fso.OpenTextFile(iResearch_Globals.iResearch_GetVersionFileName, ForReading)
    version = ts.ReadLine
    
    ts.Close
    
    lblVersion.Caption = version
    
End Sub

