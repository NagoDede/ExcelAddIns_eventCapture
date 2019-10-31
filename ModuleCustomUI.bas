Attribute VB_Name = "Module1"
Option Explicit

'Id of the Center Toggle btn
Const CENTER_TOGGLE_ID = "CenterSelectiontgl"

'The applicationEvt is used to cacth the event in the Workbook and worksheet
'The object is initialized by the LoadRibbon
Dim ApplicationEvt As ApplicationEventClass

'myRibbon is used to store the Ribbon.
Dim myRibbon As IRibbonUI

'Use to identify if a cell was formatted in CenteredSelection
Dim isCenteredSelection As Boolean
 
'******************************************************************************
' This is a callback function, used to Load the Ribbon in the interal variable.
' The sub is also used to initialize the relevant variable
' Note:
'   Sub is called in the CustomUI definition:
'   <customUI xmlns="http://schemas.microsoft.com/office/2006/01/customui"
'    onLoad="LoadRibbon" >
'******************************************************************************
Sub LoadRibbon(ribbon As IRibbonUI)
    'load the ribbon. Usefull to retrieve the toggle button and change its status
    Set myRibbon = ribbon
    ' Create the application evt and set the current activesheet
    ' By this way, even if an event sheetactive is not triggered,
    ' the macro can be executed
    Set ApplicationEvt = New ApplicationEventClass
    Call ApplicationEvt.setExcelWsh(Application.ActiveSheet)
End Sub

'******************************************************************************
' Callback function to updated the status of the CENTER_TOGGLE_ID toggle btn.
' It is called by calling myRibbon.InvalidateControl(CENTER_TOGGLE_ID)
Sub CenterAcrossSelection_Pressed(control As IRibbonControl, ByRef returnedVal)
    returnedVal = isCenteredSelection
End Sub

'******************************************************************************\
' Callback function of the CENTER_TOGGLE_ID toggle btn.
' Set the horizontal alignment of the selected cells to CenterAcrossSelection.
' If the button was pressed before, a click again will remove the centering.
' Function is call by the Ribbon button.
' Button is defined in the XML part of the xlam file.
'******************************************************************************
Sub CenterAcrossSelection(control As IRibbonControl, pressed As Boolean)
    With Selection
        If Not pressed Then
            Call RemoveCenterAcrossSelection(control)
        Else
            .HorizontalAlignment = xlCenterAcrossSelection
            .VerticalAlignment = xlCenter
            .WrapText = True
            .Orientation = 0
            .AddIndent = False
            .IndentLevel = 0
            .ShrinkToFit = False
            .ReadingOrder = xlContext
        End If
    End With
End Sub

'******************************************************************************
' this is a callback function, called when there is a change in a
' worksheet selection.
' If multiple cells are selected and one cell (at least) is not xlCenterAcrossSelection, then
' we will switch the CENTER_TOGGLE_ID to unactive.
' The CENTER_TOGGLE_ID will be active only if all the selected cells have the
' HorizontalAlignment defined to xlCenterAcrossSelection.
'
' The Definition is set in this module instead of the applicationEventClass in
' order to focus the applicationEvtClass only on the management of the events.
'******************************************************************************
Public Sub CB_SelectionChange(Target As Range)

    If Target Is Nothing Then
        Exit Sub
    End If

    Dim c As Excel.Range
    
    isCenteredSelection = True
    For Each c In Target
        If c.HorizontalAlignment <> xlCenterAcrossSelection Then
            isCenteredSelection = False
            Exit For
        End If
    Next c
    
    'This trigger the CenterAcrossSelection_Pressed call back function
    Call myRibbon.InvalidateControl(CENTER_TOGGLE_ID)

End Sub

'******************************************************************************
' Remove the xlCenterAcrossSelection Horizontal alignement of the selected cells.
'******************************************************************************
Sub RemoveCenterAcrossSelection(control As IRibbonControl)

    If TypeName(Selection) = "Range" Then
        Dim r As Excel.Range 'range
        Dim c As Excel.Range 'working cell
        Set r = Selection
        
        For Each c In r.Cells
           If c.HorizontalAlignment = xlCenterAcrossSelection Then
                     c.HorizontalAlignment = xlLeft
           End If
        Next c
    End If
End Sub

'******************************************************************************
' The following callbacks Sub change the patternColor to represent the severity
' or the applicability.
' Note: Acronyms Definition:
'       - NA: Not Appicable
'       - NSE: No Safety Effect
' Also Refer to the CustomUI definition.
'******************************************************************************
Sub setNA(control As IRibbonControl)
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = RGB(217, 217, 217)
    End With
End Sub

Sub setNSE(control As IRibbonControl)
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = RGB(183, 222, 232)
    End With
    
End Sub

Sub setMIN(control As IRibbonControl)
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = RGB(216, 228, 188)
    End With
End Sub

Sub setMAJ(control As IRibbonControl)
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = RGB(255, 255, 153)
    End With
End Sub

Sub setHAZ(control As IRibbonControl)
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = RGB(252, 213, 180)
    End With
End Sub

Sub setCAT(control As IRibbonControl)
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = RGB(230, 184, 183)
    End With
End Sub
