Attribute VB_Name = "PivotFormat"

Sub PivotFormat()
Attribute PivotFormat.VB_Description = "1) unclick autofit columns on update\n2) Display: classic Pivot Table\n3) unclick command and collapse buttons\n4) subtotals: remove   \n5)  report layout: show in tabular form\n6) report layout: repeat rows"
Attribute PivotFormat.VB_ProcData.VB_Invoke_Func = "q\n14"
'
' PivotFormat Macro
' 1) unclick autofit columns on update
' 2) Display: classic Pivot Table
' 3) unclick command and collapse buttons
' 4) subtotals: remove
' 5) report layout: show in tabular form
' 6) report layout: repeat rows
'
' Keyboard Shortcut: Ctrl+q
'
    Dim PvtTbl As PivotTable
    Dim PvtFld As PivotField
    
    For Each PvtTbl In Application.ActiveSheet.PivotTables
        With PvtTbl
            ' (1) No autoformat on update:
            .HasAutoFormat = False
            ' (2) classic display:
            .InGridDropZones = True
            ' (3) remove expand/collapse buttons:
            .ShowDrillIndicators = False
            '(4) remove subtotals:
            On Error Resume Next
            For Each PvtFld In .PivotFields
                PvtFld.Subtotals(1) = True
                PvtFld.Subtotals(1) = False
            Next PvtFld
            ' (5) Set tabluar layout:
            .RowAxisLayout xlTabularRow
            ' (6) Repeat labels:
            .RepeatAllLabels xlRepeatLabels
        End With
        
     Next
     
End Sub
