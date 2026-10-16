Attribute VB_Name = "modErrorHandler"
Option Compare Database
Option Explicit

' ============================================================================
' modErrorHandler - central error logging and user messages.
'
' Standard pattern used in every procedure:
'
'     Public Sub DoSomething()
'         On Error GoTo EH
'         ...
'     ExitHere:
'         Exit Sub
'     EH:
'         LogError "modX.DoSomething", Err.Number, Err.Description
'         Resume ExitHere
'     End Sub
'
' Err.Number / Err.Description are passed in explicitly because any
' "On Error" statement inside LogError would otherwise clear them.
' ============================================================================

Public Sub LogError(ByVal source As String, ByVal errNumber As Long, _
                    ByVal errDescription As String, Optional ByVal showUser As Boolean = True)
    On Error Resume Next

    Dim rs As DAO.Recordset
    Set rs = CurrentDb.OpenRecordset("tblErrorLog", dbOpenDynaset, dbAppendOnly)
    rs.AddNew
    rs!ErrNumber = errNumber
    rs!ErrDescription = Left$(errDescription, 4000)
    rs!ErrSource = Left$(source, 255)
    rs!UserName = Left$(Environ$("USERNAME"), 100)
    rs.Update
    rs.Close

    Debug.Print Now; " ERROR "; errNumber; " in "; source; ": "; errDescription

    If showUser Then
        MsgBox "Something went wrong." & vbCrLf & vbCrLf & _
               "Error " & errNumber & ": " & errDescription & vbCrLf & _
               "Location: " & source & vbCrLf & vbCrLf & _
               "The error has been written to tblErrorLog.", _
               vbExclamation, APP_TITLE
    End If
End Sub

Public Sub ShowInfo(ByVal message As String)
    MsgBox message, vbInformation, APP_TITLE
End Sub

Public Sub ShowWarning(ByVal message As String)
    MsgBox message, vbExclamation, APP_TITLE
End Sub

Public Function Confirm(ByVal message As String) As Boolean
    Confirm = (MsgBox(message, vbQuestion + vbYesNo + vbDefaultButton2, APP_TITLE) = vbYes)
End Function
