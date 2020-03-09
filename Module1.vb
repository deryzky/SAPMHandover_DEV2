Imports Teradata.Client.Provider

Module Module1
    Public Function Convert_Date_Str2Int(Date_Str As String) As String

        Select Case Date_Str
            Case "January"
                Convert_Date_Str2Int = "01"
            Case "February"
                Convert_Date_Str2Int = "02"
            Case "March"
                Convert_Date_Str2Int = "03"
            Case "April"
                Convert_Date_Str2Int = "04"
            Case "May"
                Convert_Date_Str2Int = "05"
            Case "June"
                Convert_Date_Str2Int = "06"
            Case "July"
                Convert_Date_Str2Int = "07"
            Case "August"
                Convert_Date_Str2Int = "08"
            Case "September"
                Convert_Date_Str2Int = "09"
            Case "October"
                Convert_Date_Str2Int = "10"
            Case "November"
                Convert_Date_Str2Int = "11"
            Case "December"
                Convert_Date_Str2Int = "12"
        End Select
        'Convert_Date_Str2Int = ""
    End Function

    Public Sub gpCompressFileToZip(ByRef sLokasiWinrar As String, ByRef sNamaFileAsli As String, ByRef sPassword As String, ByRef sNamaFileZip As String)
        Shell(sLokasiWinrar & " a -p" & sPassword & " " & sNamaFileZip & " " & sNamaFileAsli)
    End Sub

End Module
