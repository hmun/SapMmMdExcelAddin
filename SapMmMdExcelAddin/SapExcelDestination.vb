' Copyright 2016-2019 Hermann Mundprecht
' This file is licensed under the terms of the license 'CC BY 4.0'. 
' For a human readable version of the license, see https://creativecommons.org/licenses/by/4.0/

Public Class SapExcelDestination
    Private Shared ReadOnly log As log4net.ILog = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType)

    Public Function GetExcelDestinations(pWSname As String) As SAPLogon.ConParameter
        Dim conParameter As New SAPLogon.ConParameter
        Dim i As Integer
        Dim j As Integer

        Dim aPws As Excel.Worksheet
        Dim aWB As Excel.Workbook
        aWB = Globals.SapMmMdExcelAddin.Application.ActiveWorkbook
        Try
            aPws = aWB.Worksheets(pWSname)
        Catch Exc As System.Exception
            MsgBox("No " & pWSname & " Sheet in current workbook", MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAP Acc")
            GetExcelDestinations = conParameter
            Exit Function
        End Try
        i = 2
        Do Until CStr(aPws.Cells(2, i).value) = ""
            For j = 2 To 13
                If CStr(aPws.Cells(j, i).value) <> "" Then
                    log.Debug("ExcelAddOrChangeDestination - conParameter.addConValue iD=" & CStr(i - 2) & " Field=" & CStr(aPws.Cells(j, 1).value) & " Value=" & CStr(aPws.Cells(j, i).value))
                    conParameter.addConValue(CStr(i - 2), CStr(aPws.Cells(j, 1).value), CStr(aPws.Cells(j, i).value))
                End If
            Next
            i += 1
        Loop
        GetExcelDestinations = conParameter
    End Function

End Class
