Attribute VB_Name = "mProcFunc"
'***********************************************************************************
'Determine if the App.Path has a trailing "\". If not add one.
'***********************************************************************************
Public Function ftnAppPath(sAppPath As String) As String

    If Right(sAppPath$, Len(sAppPath$) - 1) <> "\" Then
        ftnAppPath$ = sAppPath$ & "\"
    End If

End Function

