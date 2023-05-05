' Copyright 2016-2019 Hermann Mundprecht
' This file is licensed under the terms of the license 'CC BY 4.0'. 
' For a human readable version of the license, see https://creativecommons.org/licenses/by/4.0/

Imports System.Reflection
Imports System.Diagnostics

Public Class SapGeneral
    Const cVersion As String = "1.0.1.2"
    Const cAssemblyName As String = "SapLtpExcelAddIn"
    Private Shared ReadOnly log As log4net.ILog = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType)
    Private _version As String
    Private _assemblyname As String

    Public Sub New()
        Try
            Dim assembly As Assembly
            Dim fileVersionInfo As FileVersionInfo
            log.Debug("checkVersion - " & "reading assembly versions")
            assembly = System.Reflection.Assembly.GetExecutingAssembly()
            fileVersionInfo = System.Diagnostics.FileVersionInfo.GetVersionInfo(assembly.Location)
            _version = fileVersionInfo.ProductVersion
            Dim assemblyName As System.Reflection.AssemblyName = assembly.GetName()
            _assemblyname = assemblyName.Name
        Catch Exc As System.Exception
            log.Debug("checkVersion - " & "failed to read assembly information using default")
            _version = cVersion
            _assemblyname = cAssemblyName
        End Try
    End Sub

    Public Function checkVersion() As Integer
        Dim aCws As Excel.Worksheet
        Dim aWB As Excel.Workbook
        Dim aFromVersion As String
        Dim aToVersion As String

        aWB = Globals.SapMmMdExcelAddin.Application.ActiveWorkbook
        Try
            aCws = aWB.Worksheets("SAP-Con")
        Catch Exc As System.Exception
            MsgBox("No SAP-Con Sheet in current workbook. Check if the current workbook is a valid SAP Accounting Template",
                   MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SapGeneral")
            checkVersion = False
            log.Error("checkVersion - Exception=" & Exc.ToString)
            Exit Function
        End Try
        log.Debug("checkVersion - " & "reading Versions")
        aFromVersion = aCws.Cells(15, 2).Value
        log.Debug("checkVersion - " & "aFromVersion=" & CStr(aFromVersion))
        aToVersion = aCws.Cells(16, 2).Value
        log.Debug("checkVersion - " & "aToVersion=" & CStr(aToVersion))

        log.Debug("checkVersion - " & "aVersion=" & CStr(_version))
        If _version > aToVersion Or _version < aFromVersion Then
            ' try Publish Version
            log.Debug("checkVersion - " & "version invalid")
            MsgBox("The Version of the Excel-Template is not valid for this Add-In. Please use a Template that is valid for version " & _version,
                   MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SapGeneral")
            checkVersion = False
            Exit Function
        End If

        log.Debug("checkVersion - " & "version OK")
        checkVersion = True
    End Function

    Public Function CheckVersionInSAP(ByRef pSapCon As SapCon) As Integer
        Dim aSapCon As Object = Nothing
        Dim aRet As Boolean
        Try
            aRet = pSapCon.getSapCon(aSapCon)
            If aRet Then
                Dim aSapVersion As New SAPLogon.SAPVersion(aSapCon)
                CheckVersionInSAP = aSapVersion.CheckVersionInSAP(_assemblyname, _version)
            Else
                CheckVersionInSAP = False
            End If
        Catch ex As SystemException
            CheckVersionInSAP = False
            log.Warn("CheckVersionInSAP - " & ex.ToString)
        End Try
    End Function

End Class
