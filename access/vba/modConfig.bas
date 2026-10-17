Attribute VB_Name = "modConfig"
Option Compare Database
Option Explicit

' ============================================================================
' modConfig - application constants, Arabic column headers and event seeds.
'
' VBA source files are stored in the Windows ANSI code page, so Arabic text
' cannot be typed into string literals safely (it turns into "????" on any
' PC whose system locale is not Arabic). Every Arabic string is therefore
' written as \uXXXX escapes and decoded at run time by UnEsc().
' ============================================================================

Public Const APP_TITLE As String = "Inventory Tracker"

Public Const IMPORT_AUTO As String = "Auto"
Public Const IMPORT_MASTER As String = "Master"
Public Const IMPORT_STATUS As String = "Status"

' How many rows at the top of a sheet are searched for the header row.
Public Const HEADER_SEARCH_ROWS As Long = 15

' Logical column ids used by the importer.
Public Const COL_WAYBILL As String = "Waybill"
Public Const COL_REFERENCE As String = "ReferenceNo"
Public Const COL_CONTENTS As String = "Contents"
Public Const COL_SECONDARY_REF As String = "SecondaryRef"
Public Const COL_CITY As String = "CityName"
Public Const COL_LAST_EVENT As String = "LastEvent"
Public Const COL_CONSIGNEE As String = "ConsigneeName"

' ---------------------------------------------------------------------------
' Header aliases: every logical column can be matched by several header
' texts. Matching is done on NormKey() so hamza / alef-maqsura / taa-marbuta
' variants and extra spaces do not matter.
' ---------------------------------------------------------------------------
Public Function HeaderAliases(ByVal logicalCol As String) As Variant
    Select Case logicalCol
        Case COL_WAYBILL
            ' "bolisat al-shahn" (master file) / "raqm al-tatabbu" (status file)
            HeaderAliases = Array( _
                UnEsc("\u0628\u0648\u0644\u064A\u0635\u0629 \u0627\u0644\u0634\u062D\u0646"), _
                UnEsc("\u0631\u0642\u0645 \u0627\u0644\u062A\u062A\u0628\u0639"), _
                UnEsc("\u0631\u0642\u0645 \u0627\u0644\u0628\u0648\u0644\u064A\u0635\u0629"), _
                "Waybill", "AWB", "Tracking Number", "Tracking No")
        Case COL_REFERENCE
            ' "raqm al-marja"
            HeaderAliases = Array( _
                UnEsc("\u0631\u0642\u0645 \u0627\u0644\u0645\u0631\u062C\u0639"), _
                "Reference", "Reference No", "Ref")
        Case COL_CONTENTS
            ' "wasf muhtawayat al-shuhna"
            HeaderAliases = Array( _
                UnEsc("\u0648\u0635\u0641 \u0645\u062D\u062A\u0648\u064A\u0627\u062A \u0627\u0644\u0634\u062D\u0646\u0629"), _
                "Contents", "Description", "Shipment Contents")
        Case COL_SECONDARY_REF
            ' "raqm al-marja al-thanawi"
            HeaderAliases = Array( _
                UnEsc("\u0631\u0642\u0645 \u0627\u0644\u0645\u0631\u062C\u0639 \u0627\u0644\u062B\u0627\u0646\u0648\u064A"), _
                "Secondary Reference", "Secondary Ref")
        Case COL_CITY
            ' "madinat al-wijha"
            HeaderAliases = Array( _
                UnEsc("\u0645\u062F\u064A\u0646\u0629 \u0627\u0644\u0648\u062C\u0647\u0629"), _
                "Destination City", "City")
        Case COL_LAST_EVENT
            ' "akhir hadath"
            HeaderAliases = Array( _
                UnEsc("\u0622\u062E\u0631 \u062D\u062F\u062B"), _
                "Last Event", "Status")
        Case COL_CONSIGNEE
            ' "ism al-mursal ilayh"
            HeaderAliases = Array( _
                UnEsc("\u0627\u0633\u0645 \u0627\u0644\u0645\u0631\u0633\u0644 \u0625\u0644\u064A\u0647"), _
                "Consignee", "Consignee Name", "Receiver")
        Case Else
            HeaderAliases = Array()
    End Select
End Function

' ---------------------------------------------------------------------------
' Seed events. Element layout: Array(EventName, IsInWarehouse, SortOrder)
'   "jahiza lil-iltiqat"   (ready for pickup)       -> IN warehouse
'   "istilam min al-amil"  (received from customer) -> IN warehouse
'   "al-taslim al-niha'i"  (final delivery)         -> OUT
' FIXED RULE: only the events flagged True here are IN warehouse. Every
' other event (including any new one that appears in a status file) is OUT.
' EnforceInRule re-applies this on every status import, so the flag cannot
' drift through manual edits to tblEventTypes.
' ---------------------------------------------------------------------------
Public Function SeedEvents() As Variant
    SeedEvents = Array( _
        Array(UnEsc("\u062C\u0627\u0647\u0632\u0629 \u0644\u0644\u0627\u0644\u062A\u0642\u0627\u0637"), True, 10), _
        Array(UnEsc("\u0627\u0633\u062A\u0644\u0627\u0645 \u0645\u0646 \u0627\u0644\u0639\u0645\u064A\u0644"), True, 20), _
        Array(UnEsc("\u0627\u0644\u062A\u0633\u0644\u064A\u0645 \u0627\u0644\u0646\u0647\u0627\u0626\u064A"), False, 90))
End Function

' Decodes \uXXXX escapes; every other character is copied unchanged.
Public Function UnEsc(ByVal s As String) As String
    Dim i As Long, out As String
    i = 1
    Do While i <= Len(s)
        If Mid$(s, i, 2) = "\u" And i + 5 <= Len(s) Then
            out = out & ChrW$(CLng("&H" & Mid$(s, i + 2, 4)))
            i = i + 6
        Else
            out = out & Mid$(s, i, 1)
            i = i + 1
        End If
    Loop
    UnEsc = out
End Function
