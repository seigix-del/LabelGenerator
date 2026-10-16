Attribute VB_Name = "modNavigation"
Option Compare Database
Option Explicit

' ============================================================================
' modNavigation - main menu / dashboard, form navigation and button handlers.
'
' Buttons call these as expressions in their event properties, e.g.
' OnClick = "=Nav_OpenImport()". That keeps the forms free of code-behind
' (except where an event must be cancelled, see modBuildDatabase).
' ============================================================================

Public Const FRM_MAIN As String = "frmMainMenu"
Public Const FRM_IMPORT As String = "frmImport"
Public Const FRM_SHIPMENTS As String = "frmShipments"
Public Const FRM_HISTORY As String = "frmShipmentHistory"
Public Const FRM_EVENTS As String = "frmEventTypes"
Public Const FRM_UNMATCHED As String = "frmUnmatched"
Public Const FRM_BATCHES As String = "frmImportHistory"

Public Const LOC_ALL As String = "All"
Public Const LOC_IN As String = "In warehouse"
Public Const LOC_OUT As String = "Out of warehouse"

' ---------------------------------------------------------------------------
' Generic helpers
' ---------------------------------------------------------------------------

Public Function IsFormLoaded(ByVal formName As String) As Boolean
    On Error Resume Next
    IsFormLoaded = CurrentProject.AllForms(formName).IsLoaded
End Function

Public Sub OpenFormSafe(ByVal formName As String, Optional ByVal whereCondition As String = "")
    On Error GoTo EH
    If IsFormLoaded(formName) And Len(whereCondition) = 0 Then
        DoCmd.SelectObject acForm, formName
        Forms(formName).Requery
    Else
        If IsFormLoaded(formName) Then DoCmd.Close acForm, formName, acSaveNo
        DoCmd.OpenForm formName, acNormal, , whereCondition
    End If
ExitHere:
    Exit Sub
EH:
    LogError "modNavigation.OpenFormSafe(" & formName & ")", Err.Number, Err.Description
    Resume ExitHere
End Sub

Public Sub CloseFormSafe(ByVal formName As String)
    On Error Resume Next
    If IsFormLoaded(formName) Then DoCmd.Close acForm, formName, acSaveNo
End Sub

' ---------------------------------------------------------------------------
' Dashboard KPIs (used as control sources on frmMainMenu: "=Kpi_InWarehouse()")
' ---------------------------------------------------------------------------

Public Function Kpi_Total() As Long
    Kpi_Total = Nz(DSum("Quantity", "tblOrderDetails"), 0)
End Function

Public Function Kpi_InWarehouse() As Long
    Kpi_InWarehouse = Nz(DSum("Quantity", "tblOrderDetails", "InWarehouse=True"), 0)
End Function

Public Function Kpi_OutOfWarehouse() As Long
    Kpi_OutOfWarehouse = Nz(DSum("Quantity", "tblOrderDetails", "InWarehouse=False"), 0)
End Function

Public Function Kpi_InPercent() As String
    Dim t As Long
    t = Kpi_Total()
    If t = 0 Then
        Kpi_InPercent = "-"
    Else
        Kpi_InPercent = Format$(Kpi_InWarehouse() / t, "0.0%")
    End If
End Function

Public Function Kpi_LastImport() As String
    Dim id As Variant
    id = DMax("BatchID", "tblImportBatches")
    If IsNull(id) Then
        Kpi_LastImport = "No imports yet"
    Else
        Kpi_LastImport = DLookup("ImportType", "tblImportBatches", "BatchID=" & id) & " - " & _
                         Format$(DLookup("ImportedOn", "tblImportBatches", "BatchID=" & id), "yyyy-mm-dd hh:nn")
    End If
End Function

Public Function Kpi_LastStatusChanged() As Long
    Kpi_LastStatusChanged = Nz(DLookup("RowsUpdated", "qryRpt_LastStatusBatch"), 0)
End Function

Public Function Kpi_LastStatusUnmatched() As Long
    Kpi_LastStatusUnmatched = Nz(DLookup("RowsUnmatched", "qryRpt_LastStatusBatch"), 0)
End Function

Public Function Dash_Refresh() As Boolean
    On Error Resume Next
    If IsFormLoaded(FRM_MAIN) Then
        Forms(FRM_MAIN).Recalc
        Forms(FRM_MAIN).Controls("lstByEvent").Requery
        Forms(FRM_MAIN).Controls("lstByProduct").Requery
    End If
    Dash_Refresh = True
End Function

' ---------------------------------------------------------------------------
' Navigation buttons
' ---------------------------------------------------------------------------

Public Function Nav_OpenImport() As Boolean
    OpenFormSafe FRM_IMPORT
End Function

Public Function Nav_OpenShipments() As Boolean
    OpenShipmentsFiltered LOC_ALL
End Function

Public Function Nav_OpenInWarehouse() As Boolean
    OpenShipmentsFiltered LOC_IN
End Function

Public Function Nav_OpenOutOfWarehouse() As Boolean
    OpenShipmentsFiltered LOC_OUT
End Function

Public Function Nav_OpenEventTypes() As Boolean
    OpenFormSafe FRM_EVENTS
End Function

Public Function Nav_OpenUnmatched() As Boolean
    OpenFormSafe FRM_UNMATCHED
End Function

Public Function Nav_OpenImportHistory() As Boolean
    OpenFormSafe FRM_BATCHES
End Function

' Double-click on a shipment row -> its status history.
Public Function Nav_OpenHistoryForCurrent() As Boolean
    On Error GoTo EH
    Dim id As Variant
    id = Screen.ActiveForm!DetailID
    If IsNull(id) Then Exit Function
    OpenFormSafe FRM_HISTORY, "DetailID=" & id
ExitHere:
    Exit Function
EH:
    LogError "modNavigation.Nav_OpenHistoryForCurrent", Err.Number, Err.Description
    Resume ExitHere
End Function

' "Close" button on every child form.
Public Function Nav_CloseActive() As Boolean
    On Error Resume Next
    Dim n As String
    n = Screen.ActiveForm.Name
    If n <> FRM_MAIN Then DoCmd.Close acForm, n, acSaveNo
    If IsFormLoaded(FRM_MAIN) Then
        DoCmd.SelectObject acForm, FRM_MAIN
        Dash_Refresh
    End If
End Function

Public Function Nav_Exit() As Boolean
    If Confirm("Close the inventory database?") Then DoCmd.Quit acQuitSaveAll
End Function

' ---------------------------------------------------------------------------
' frmImport
' ---------------------------------------------------------------------------

Public Function Imp_Browse() As Boolean
    On Error GoTo EH
    Dim p As String
    p = PickExcelFile("Select the inventory or status Excel file")
    If Len(p) > 0 Then
        Forms(FRM_IMPORT)!txtFilePath = p
        Forms(FRM_IMPORT)!txtResult = Null
    End If
ExitHere:
    Exit Function
EH:
    LogError "modNavigation.Imp_Browse", Err.Number, Err.Description
    Resume ExitHere
End Function

Public Function Imp_Run() As Boolean
    On Error GoTo EH
    Dim f As Form, summary As String, errMsg As String, kind As String

    Set f = Forms(FRM_IMPORT)
    kind = Nz(f!cboImportType, IMPORT_AUTO)

    If Not ValidateImportFile(Nz(f!txtFilePath, ""), errMsg) Then
        ShowWarning errMsg
        f!txtFilePath.SetFocus
        Exit Function
    End If

    f!txtResult = "Importing, please wait..."
    DoEvents

    If ImportExcelFile(f!txtFilePath, kind, summary) Then
        f!txtResult = summary
        Dash_Refresh
        If IsFormLoaded(FRM_SHIPMENTS) Then Forms(FRM_SHIPMENTS).Requery
    Else
        f!txtResult = summary
    End If
ExitHere:
    Exit Function
EH:
    LogError "modNavigation.Imp_Run", Err.Number, Err.Description
    Resume ExitHere
End Function

' ---------------------------------------------------------------------------
' frmShipments - search & IN/OUT filter
' ---------------------------------------------------------------------------

Private Sub OpenShipmentsFiltered(ByVal location As String)
    OpenFormSafe FRM_SHIPMENTS
    If IsFormLoaded(FRM_SHIPMENTS) Then
        Forms(FRM_SHIPMENTS)!cboLocation = location
        Shp_ApplyFilter
    End If
End Sub

Public Function Shp_ApplyFilter() As Boolean
    On Error GoTo EH
    Dim f As Form, crit As String, s As String, loc As String

    Set f = Forms(FRM_SHIPMENTS)
    s = NormText(f!txtSearch)
    loc = Nz(f!cboLocation, LOC_ALL)

    If Len(s) > 0 Then
        s = "'*" & Replace(s, "'", "''") & "*'"
        crit = "(Waybill Like " & s & " OR ReferenceNo Like " & s & _
               " OR SecondaryRef Like " & s & " OR ProductDescription Like " & s & _
               " OR CustomerName Like " & s & " OR CityName Like " & s & _
               " OR EventName Like " & s & ")"
    End If

    Select Case loc
        Case LOC_IN:  crit = crit & IIf(Len(crit) > 0, " AND ", "") & "InWarehouse=True"
        Case LOC_OUT: crit = crit & IIf(Len(crit) > 0, " AND ", "") & "InWarehouse=False"
    End Select

    If Len(crit) = 0 Then
        f.FilterOn = False
        f.Filter = ""
    Else
        f.Filter = crit
        f.FilterOn = True
    End If
    f!txtCount.Requery
ExitHere:
    Exit Function
EH:
    LogError "modNavigation.Shp_ApplyFilter", Err.Number, Err.Description
    Resume ExitHere
End Function

Public Function Shp_ClearFilter() As Boolean
    On Error Resume Next
    With Forms(FRM_SHIPMENTS)
        !txtSearch = Null
        !cboLocation = LOC_ALL
    End With
    Shp_ApplyFilter
End Function

' ---------------------------------------------------------------------------
' frmEventTypes
' ---------------------------------------------------------------------------

' AfterUpdate of an event rule: push the new IN/OUT flag to every waybill
' that currently sits on that event.
Public Function Evt_AfterUpdate() As Boolean
    Dim n As Long
    n = RecalcInWarehouse()
    If n > 0 Then ShowInfo n & " waybill(s) moved between IN and OUT after this rule change."
    Dash_Refresh
End Function

' ---------------------------------------------------------------------------
' Exports
' ---------------------------------------------------------------------------

Public Function Exp_InWarehouse() As Boolean
    ExportQuery "qryRpt_InWarehouse", "InWarehouse_" & Format$(Date, "yyyymmdd")
End Function

Public Function Exp_OutOfWarehouse() As Boolean
    ExportQuery "qryRpt_OutOfWarehouse", "OutOfWarehouse_" & Format$(Date, "yyyymmdd")
End Function

Public Function Exp_Unmatched() As Boolean
    ExportQuery "qryRpt_UnmatchedLastBatch", "Unmatched_" & Format$(Date, "yyyymmdd")
End Function

Private Sub ExportQuery(ByVal queryName As String, ByVal defaultName As String)
    On Error GoTo EH
    Dim p As String
    p = PickSavePath(defaultName & ".xlsx")
    If Len(p) = 0 Then Exit Sub
    DoCmd.OutputTo acOutputQuery, queryName, acFormatXLSX, p, True
ExitHere:
    Exit Sub
EH:
    LogError "modNavigation.ExportQuery(" & queryName & ")", Err.Number, Err.Description
    Resume ExitHere
End Sub
