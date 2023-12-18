' Copyright 2023 Hermann Mundprecht
' This file is licensed under the terms of the license 'CC BY 4.0'. 
' For a human readable version of the license, see https://creativecommons.org/licenses/by/4.0/

Public Class SapMmMdRibbon_Routing
    Private Shared ReadOnly log As log4net.ILog = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType)
    Public Function getGenParameters(ByRef pPar As SAPCommon.TStr) As Integer
        Dim aPws As Excel.Worksheet
        Dim aWB As Excel.Workbook
        Dim aKey As String
        Dim aName As String
        Dim i As Integer
        log.Debug("SapMmMdRibbon_Routing getGenParametrs - " & "reading Parameter")
        aWB = Globals.SapMmMdExcelAddin.Application.ActiveWorkbook
        Try
            aPws = aWB.Worksheets("Parameter")
        Catch Exc As System.Exception
            MsgBox("No Parameter Sheet in current workbook. Check if the current workbook is a valid SAP MM Md Template",
                   MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap MM Md")
            getGenParameters = False
            Exit Function
        End Try
        aName = "SAPMmMdMaterial"
        aKey = CStr(aPws.Cells(1, 1).Value)
        If aKey <> aName Then
            MsgBox("Cell A1 of the parameter sheet does not contain the key " & aName & ". Check if the current workbook is a valid SAP MM Md Template",
                   MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap MM Md")
            getGenParameters = False
            Exit Function
        End If
        i = 2
        pPar = New SAPCommon.TStr
        Do While CStr(aPws.Cells(i, 2).value) <> "" Or CStr(aPws.Cells(i, 2).value) <> ""
            pPar.add(CStr(aPws.Cells(i, 2).value), CStr(aPws.Cells(i, 4).value), pFORMAT:=CStr(aPws.Cells(i, 3).value))
            i += 1
        Loop
        getGenParameters = True
    End Function

    Private Function getIntParameters(ByRef pIntPar As SAPCommon.TStr) As Integer
        Dim aPws As Excel.Worksheet
        Dim aWB As Excel.Workbook
        Dim i As Integer

        log.Debug("getIntParameters - " & "reading Parameter")
        aWB = Globals.SapMmMdExcelAddin.Application.ActiveWorkbook
        Try
            aPws = aWB.Worksheets("Parameter_Int")
        Catch Exc As System.Exception
            MsgBox("No Parameter_Int Sheet in current workbook. Check if the current workbook is a valid SAP MM Md Template",
                   MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap MM Md")
            getIntParameters = False
            Exit Function
        End Try
        i = 2
        pIntPar = New SAPCommon.TStr
        Do
            pIntPar.add(CStr(aPws.Cells(i, 2).value), CStr(aPws.Cells(i, 3).value))
            i += 1
        Loop While CStr(aPws.Cells(i, 2).value) <> "" Or CStr(aPws.Cells(i, 2).value) <> ""
        ' no obligatory parameters check - we should know what we are doing
        getIntParameters = True
    End Function

    Public Sub Maintain(ByRef pSapCon As SapCon, Optional pMode As String = "create")
        Dim aSAPRouting As New SAPRouting(pSapCon)

        Dim aPar As New SAPCommon.TStr
        Dim aIntPar As New SAPCommon.TStr

        ' get general parameters
        If getGenParameters(aPar) = False Then
            Exit Sub
        End If
        ' get internal parameters
        If Not getIntParameters(aIntPar) Then
            Exit Sub
        End If

        Dim jMax As UInt64 = 0
        Dim aRouLOff As Integer = If(aIntPar.value("LOFF", "ROU_DATA") <> "", CInt(aIntPar.value("LOFF", "ROU_DATA")), 4)
        Dim aHdrLOff As Integer = If(aIntPar.value("LOFF", "ROU_HOFF") <> "", CInt(aIntPar.value("LOFF", "ROU_HOFF")), aRouLOff - 3)
        Dim aRouWsName As String = If(aIntPar.value("WS", "ROU_DATA") <> "", aIntPar.value("WS", "ROU_DATA"), "Data")
        Dim aRouWs As Excel.Worksheet
        Dim aMsgClmn As String = If(aIntPar.value("COL", "DATAMSG") <> "", aIntPar.value("COL", "DATAMSG"), "INT-MSG")
        Dim aMsgClmnNr As Integer = 0
        Dim aMatClmnNr As Integer = If(aIntPar.value("COLNR", "DATAMAT") <> "", CInt(aIntPar.value("COLNR", "DATAMAT")), 1)
        Dim aRetStr As String
        Dim aOKMsg As String = If(aIntPar.value("RET", "OKMSG") <> "", aIntPar.value("RET", "OKMSG"), "OK")

        Dim aWB As Excel.Workbook
        aWB = Globals.SapMmMdExcelAddin.Application.ActiveWorkbook
        Try
            aRouWs = aWB.Worksheets(aRouWsName)
        Catch Exc As System.Exception
            MsgBox("No " & aRouWsName & " Sheet in current workbook. Check if the current workbook is a valid SAP Materal Template",
                   MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAP Materal")
            Exit Sub
        End Try
        parseHeaderLine(aRouWs, jMax, aMsgClmn, aMsgClmnNr, pHdrLine:=aHdrLOff)
        Try
            log.Debug("SapMmMdRibbon_Routing.Maintain - " & "processing data - disabling events, screen update, cursor")
            Globals.SapMmMdExcelAddin.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlWait
            Globals.SapMmMdExcelAddin.Application.EnableEvents = False
            Globals.SapMmMdExcelAddin.Application.ScreenUpdating = False
            Dim i As UInt64 = aRouLOff + 1
            Dim aKey As String
            Dim aItems As New TData(aIntPar)
            Dim aTSAP_RoutingData As New TSAP_RoutingData(aPar, aIntPar, aSAPRouting, pMode)
            Do
                If Left(CStr(aRouWs.Cells(i, aMsgClmnNr).Value), Len(aOKMsg)) <> aOKMsg Then
                    aKey = CStr(i)
                    aRetStr = ""
                    ' read DATA
                    aItems.ws_parse_line_simple(aRouWs, aRouLOff, i, jMax)
                    If CStr(aRouWs.Cells(i, aMatClmnNr).Value) <> CStr(aRouWs.Cells(i + 1, aMatClmnNr).Value) Then
                        If aTSAP_RoutingData.fillHeader(aItems) And aTSAP_RoutingData.fillData(aItems) Then
                            If pMode = "create" Then
                                log.Debug("SapMmMdRibbon_Routing.Maintain - " & "calling aSAPRouting.Create")
                                aRetStr = aSAPRouting.Create(aTSAP_RoutingData, pOKMsg:=aOKMsg)
                                log.Debug("SapMmMdRibbon_Routing.Maintain - " & "aSAPRouting.Create returned, aRetStr=" & aRetStr)
                            ElseIf pMode = "change" Then
                                log.Debug("SapMmMdRibbon_Routing.Maintain - " & "calling aSAPRouting.Change")
                                aRetStr = aSAPRouting.Change(aTSAP_RoutingData, pOKMsg:=aOKMsg)
                                log.Debug("SapMmMdRibbon_Routing.Maintain - " & "aSAPRouting.Change returned, aRetStr=" & aRetStr)
                            End If
                            For Each aKey In aItems.aTDataDic.Keys
                                aRouWs.Cells(CInt(aKey), aMsgClmnNr) = CStr(aRetStr)
                            Next
                            aItems = New TData(aIntPar)
                            aTSAP_RoutingData = New TSAP_RoutingData(aPar, aIntPar, aSAPRouting, pMode)
                        End If
                    End If
                End If
                i += 1
            Loop While Not String.IsNullOrEmpty(CStr(aRouWs.Cells(i, 1).value))
            log.Debug("SapMmMdRibbon_Routing.Maintain - " & "all data processed - enabling events, screen update, cursor")
            Globals.SapMmMdExcelAddin.Application.EnableEvents = True
            Globals.SapMmMdExcelAddin.Application.ScreenUpdating = True
            Globals.SapMmMdExcelAddin.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault
        Catch ex As System.Exception
            Globals.SapMmMdExcelAddin.Application.EnableEvents = True
            Globals.SapMmMdExcelAddin.Application.ScreenUpdating = True
            Globals.SapMmMdExcelAddin.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault
            MsgBox("SapMmMdRibbon_Routing.Maintain failed! " & ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAP Materal AddIn")
            log.Error("SapMmMdRibbon_Routing.Maintain - " & "Exception=" & ex.ToString)
            Exit Sub
        End Try
    End Sub

    Private Sub parseHeaderLine(ByRef pWs As Excel.Worksheet, ByRef pMaxJ As Integer, Optional pMsgClmn As String = "", Optional ByRef pMsgClmnNr As Integer = 0, Optional pHdrLine As Integer = 1)
        pMaxJ = 0
        Do
            pMaxJ += 1
            If Not String.IsNullOrEmpty(pMsgClmn) And CStr(pWs.Cells(pHdrLine, pMaxJ).value) = pMsgClmn Then
                pMsgClmnNr = pMaxJ
            End If
        Loop While CStr(pWs.Cells(pHdrLine, pMaxJ + 1).value) <> ""
    End Sub

End Class
