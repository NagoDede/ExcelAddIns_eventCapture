VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ApplicationEventClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'******************************************************************************
' The code is used inside an Add-ins.
' To retrieve the events on the worksheet, we need to retrieve the Application
' and worksheet events by a different manner.
' The aim of the class is to retrieve the events.
'******************************************************************************
'To retrieve the application events
Dim WithEvents ExcelAppEvents As Application
Attribute ExcelAppEvents.VB_VarHelpID = -1
'To retrieve the worksheet events
Dim WithEvents ExcelWshEvents As Worksheet
Attribute ExcelWshEvents.VB_VarHelpID = -1

'******************************************************************************
' During the init, we get the application object
Private Sub Class_Initialize()
    Set ExcelAppEvents = Application
End Sub

'******************************************************************************
' when the active sheet change, we store the active worksheet
Private Sub ExcelAppEvents_SheetActivate(ByVal sh As Object)
    Set ExcelWshEvents = sh
End Sub

'******************************************************************************
' Provides a way to change the ExcelWorksheet.
' Usefull during init process to apply the active sheet as there
' is no application event trigger.
Public Sub setExcelWsh(sh As Worksheet)
    If Not sh Is Nothing Then
        Set ExcelWshEvents = sh
    End If
End Sub

'******************************************************************************
' When the active sheet change, we store the active worksheet
' This shoul be useless, but not sure....
Private Sub ExcelAppEvents_SheetChange(ByVal sh As Object, ByVal Target As Range)
    Set ExcelWshEvents = sh
End Sub

'******************************************************************************
' This catch the Selection change (Cell change) in the identified Worksheet.
' In order to focus the class only on the events management, we call a sub in
' the main module.
Private Sub ExcelWshEvents_SelectionChange(ByVal Target As Range)
    Call CB_SelectionChange(Target)
End Sub
