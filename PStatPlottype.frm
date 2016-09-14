VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} PStatPlottype 
   Caption         =   "Pick the type of plot for this data:"
   ClientHeight    =   1575
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4890
   OleObjectBlob   =   "PStatPlottype.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "PStatPlottype"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub axesEIvsTime_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Me.Tag = "axesEIvsTime"
    Me.Hide
End Sub

Private Sub axesEvsI_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Me.Tag = "axesEvsI"
    Me.Hide
End Sub

Private Sub axesEvslogI_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Me.Tag = "axesEvslogI"
    Me.Hide
End Sub

Private Sub axesEvsTime_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Me.Tag = "axesEvsTime"
    Me.Hide
End Sub

Private Sub axesIvsTime_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Me.Tag = "axesIvsTime"
    Me.Hide
End Sub

Private Sub btnAbortPlot_Click()
    ' Make no selection
    If axesEIvsTime = True Then axesEIvsTime = False
    If axesEvsI = True Then axesEvsI = False
    If axesEvslogI = True Then axesEvslogI = False
    If axesEvsTime = True Then axesEvsTime = False
    If axesIvsTime = True Then axesIvsTime = False
    Me.Tag = vbNullString
    ' Hide the form
    Me.Hide
End Sub

Private Sub btnSubmitPlotChoice_Click()
    ' Check for a plottype selection
    If axesEIvsTime + axesEvsI + axesEvslogI + axesEvsTime + axesIvsTime = 0 Then
        Exit Sub
    Else
        ' Store the selection
        If axesEIvsTime = True Then Me.Tag = "axesEIvsTime"
        If axesEvsI = True Then Me.Tag = "axesEvsI"
        If axesEvslogI = True Then Me.Tag = "axesEvslogI"
        If axesEvsTime = True Then Me.Tag = "axesEvsTime"
        If axesIvsTime = True Then Me.Tag = "axesIvsTime"
    End If
    Me.Hide
End Sub
