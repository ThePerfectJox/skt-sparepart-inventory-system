Attribute VB_Name = "coreConfig"

Option Explicit

Private Const DATABASE_FILE_NAME As String = "SIS.accdb"

Public Function currentUser() As String
    currentUser = Environ$("USERNAME")
End Function

Public Function currentExcelPath() As String
    currentExcelPath = ThisWorkbook.Path
End Function

Public Function SharePointToLocalPath(ByVal sharePointUrl As String) As String

    Dim localRoot As String
    Dim pos As Long

    localRoot = "C:\Users\" & currentUser() & "\OneDrive - Philip Morris International\"
    sharePointUrl = Replace(sharePointUrl, "/", "\")
    pos = InStr(1, sharePointUrl, "\Documents\", vbTextCompare)

    If pos > 0 Then
        SharePointToLocalPath = _
            localRoot & mId$(sharePointUrl, pos + Len("\Documents\"))
    Else
        SharePointToLocalPath = sharePointUrl
    End If

End Function


Public Function DatabasePath() As String
    Dim settingsRange As Range
    Dim excelPath As String
    Dim pathCol As Collection
    Dim result As String
    Dim part As String
    Dim i As Long

    Set settingsRange = _
        ThisWorkbook.Worksheets("SETTINGS").Range("L10:R10")

    If Application.CountA(settingsRange) > 0 Then
        excelPath = Replace(CStr(settingsRange.Cells(1, 1).value), """", "")
    Else
        excelPath = currentExcelPath()
    End If

    Set pathCol = coreHelper.SplitString(excelPath, "\")
    For i = 1 To pathCol.count
        part = CStr(pathCol(i))
        If i > 1 Then
            If StrComp( _
                CStr(pathCol(i - 1)), _
                "Users", _
                vbTextCompare _
            ) = 0 Then
                part = currentUser()
            End If
        End If

        If result <> "" Then
            result = result & "\"
        End If

        result = result & part

    Next i

    If InStrRev(result, "\") > 0 Then

        part = mId$(result, InStrRev(result, "\") + 1)

        If InStr(part, ".") > 0 Then
            result = Left$(result, InStrRev(result, "\"))
        Else
            result = result & "\"
        End If
    End If
    
    Debug.Print SharePointToLocalPath( _
        result & DATABASE_FILE_NAME _
    )
    
    DatabasePath = SharePointToLocalPath( _
        result & DATABASE_FILE_NAME _
    )
    
End Function

Public Function ApplicationName() As String
    ApplicationName = "Sparepart Inventory System"
End Function
