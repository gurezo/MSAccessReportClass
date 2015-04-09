Option Strict Off
Option Explicit On 

Public Class MSAccessReport

    Private oleAccess As Object

    Public Function lngPreview(ByRef DbName As String, _
                               ByRef RptName As String, _
                               ByRef ConditionStr As String) As Long

        Dim Rtn As Long

        lngPreview = 0

        Rtn = lngCreateAccessObj(DbName)
        If Rtn <> 0 Then
            lngPreview = Rtn
            Exit Function
        End If

        Try
            oleAccess.DoCmd.OpenReport(RptName, Access.AcFormView.acPreview, , ConditionStr)
            oleAccess.DoCmd.Maximize()
            oleAccess.Visible = True
        Catch ex As Exception
            lngPreview = Err.Number
        End Try

    End Function

    Public Function lngPrint(ByRef DbName As String, _
                               ByRef RptName As String, _
                               ByRef ConditionStr As String) As Long
        Dim Rtn As Long

        lngPrint = 0

        Rtn = lngCreateAccessObj(DbName)
        If Rtn <> 0 Then
            lngPrint = Rtn
            Exit Function
        End If

        Try
            oleAccess.DoCmd.OpenReport(RptName, Access.AcFormView.acNormal, , ConditionStr)
        Catch ex As Exception
            lngPrint = Err.Number
        End Try

        oleAccess.CloseCurrentDatabase()
        oleAccess = Nothing

    End Function

    Private Function lngCreateAccessObj(ByRef DbName As String) As Long

        lngCreateAccessObj = 0

        oleAccess = Nothing
        oleAccess = New Access.Application

        Try
            oleAccess.OpenCurrentDatabase(DbName, False)
        Catch ex As Exception
            lngCreateAccessObj = Err.Number
        End Try

    End Function

End Class
