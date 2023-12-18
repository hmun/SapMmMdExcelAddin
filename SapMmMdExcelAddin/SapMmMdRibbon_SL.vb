' Copyright 2022 Hermann Mundprecht, Stefan Duben
' This file is licensed under the terms of the license 'CC BY 4.0'. 
' For a human readable version of the license, see https://creativecommons.org/licenses/by/4.0/

Public Class SapMmMdRibbon_SL
    Private Shared ReadOnly log As log4net.ILog = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType)
    Public Function getGenParameters(ByRef pPar As SAPCommon.TStr) As Integer
        Dim aPws As Excel.Worksheet
        Dim aWB As Excel.Workbook
        Dim aKey As String
        Dim aName As String
        Dim i As Integer
        log.Debug("SapMmMdRibbon_SL getGenParametrs - " & "reading Parameter")
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

    Public Sub Update(ByRef pSapCon As SapCon)
        Dim aSAPSourceList As New SAPSourceList(pSapCon)

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
        Dim aSLLOff As Integer = If(aIntPar.value("LOFF", "SL_DATA") <> "", CInt(aIntPar.value("LOFF", "SL_DATA")), 4)
        Dim aSLWsName As String = If(aIntPar.value("WS", "SL_DATA") <> "", aIntPar.value("WS", "SL_DATA"), "Data")
        Dim aSLWs As Excel.Worksheet
        Dim aMsgClmn As String = If(aIntPar.value("COL", "DATAMSG") <> "", aIntPar.value("COL", "DATAMSG"), "INT-MSG")
        Dim aMsgClmnNr As Integer = 0
        Dim aMatClmnNr As Integer = If(aIntPar.value("COLNR", "DATMAT") <> "", CInt(aIntPar.value("COLNR", "DATMAT")), 1)
        Dim aRetStr As String
        Dim aOKMsg As String = If(aIntPar.value("RET", "OKMSG") <> "", aIntPar.value("RET", "OKMSG"), "OK")

        Dim aWB As Excel.Workbook
        aWB = Globals.SapMmMdExcelAddin.Application.ActiveWorkbook
        Try
            aSLWs = aWB.Worksheets(aSLWsName)
        Catch Exc As System.Exception
            MsgBox("No " & aSLWsName & " Sheet in current workbook. Check if the current workbook is a valid SAP Materal Template",
                   MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAP Materal")
            Exit Sub
        End Try
        parseHeaderLine(aSLWs, jMax, aMsgClmn, aMsgClmnNr, pHdrLine:=aSLLOff - 3)
        Try
            log.Debug("SapMmMdRibbon_SL.Update - " & "processing data - disabling events, screen update, cursor")
            Globals.SapMmMdExcelAddin.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlWait
            Globals.SapMmMdExcelAddin.Application.EnableEvents = False
            Globals.SapMmMdExcelAddin.Application.ScreenUpdating = False
            Dim i As UInt64 = aSLLOff + 1
            Dim aKey As String
            Dim aSLItems As New TData(aIntPar)
            Dim aTSAP_SourceListData As New TSAP_SourceListData(aPar, aIntPar, aSAPSourceList, "Update")
            Do
                If Left(CStr(aSLWs.Cells(i, aMsgClmnNr).Value), Len(aOKMsg)) <> aOKMsg Then
                    aKey = CStr(i)
                    aRetStr = ""
                    ' read DATA
                    aSLItems.ws_parse_line_simple(aSLWs, aSLLOff, i, jMax)
                    If CStr(aSLWs.Cells(i, aMatClmnNr).Value) <> CStr(aSLWs.Cells(i + 1, aMatClmnNr).Value) Then
                        If aTSAP_SourceListData.fillHeader(aSLItems) And aTSAP_SourceListData.fillData(aSLItems) Then
                            log.Debug("SapMmMdRibbon_SL.Update - " & "calling aSAPSourceList.Update")
                            aRetStr = aSAPSourceList.Update(aTSAP_SourceListData, pOKMsg:=aOKMsg)
                            log.Debug("SapMmMdRibbon_SL.Update - " & "aSAPSourceList.Update returned, aRetStr=" & aRetStr)
                            For Each aKey In aSLItems.aTDataDic.Keys
                                aSLWs.Cells(CInt(aKey), aMsgClmnNr) = CStr(aRetStr)
                            Next
                            aSLItems = New TData(aIntPar)
                            aTSAP_SourceListData = New TSAP_SourceListData(aPar, aIntPar, aSAPSourceList, "Update")
                        End If
                    End If
                End If
                i += 1
            Loop While Not String.IsNullOrEmpty(CStr(aSLWs.Cells(i, 1).value))
            log.Debug("SapMmMdRibbon_SL.Update - " & "all data processed - enabling events, screen update, cursor")
            Globals.SapMmMdExcelAddin.Application.EnableEvents = True
            Globals.SapMmMdExcelAddin.Application.ScreenUpdating = True
            Globals.SapMmMdExcelAddin.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault
        Catch ex As System.Exception
            Globals.SapMmMdExcelAddin.Application.EnableEvents = True
            Globals.SapMmMdExcelAddin.Application.ScreenUpdating = True
            Globals.SapMmMdExcelAddin.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault
            MsgBox("SapMmMdRibbon_SL.Update failed! " & ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAP Materal AddIn")
            log.Error("SapMmMdRibbon_SL.Update - " & "Exception=" & ex.ToString)
            Exit Sub
        End Try

    End Sub

    Public Sub Read(ByRef pSapCon As SapCon)
        Dim aSAPSourceList As New SAPSourceList(pSapCon)

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
        Dim aMLiLOff As Integer = If(aIntPar.value("LOFF", "SL_LIST") <> "", CInt(aIntPar.value("LOFF", "SL_LIST")), 4)
        Dim aSLLOff As Integer = If(aIntPar.value("LOFF", "SL_DATA") <> "", CInt(aIntPar.value("LOFF", "SL_DATA")), 4)
        Dim aMLiWsName As String = If(aIntPar.value("WS", "SL_LIST") <> "", aIntPar.value("WS", "SL_LIST"), "Material_List")
        Dim aSLWsName As String = If(aIntPar.value("WS", "SL_DATA") <> "", aIntPar.value("WS", "SL_DATA"), "Data")
        Dim aSLWs As Excel.Worksheet
        Dim aMLiWs As Excel.Worksheet
        Dim aMsgClmn As String = If(aIntPar.value("COL", "LISTMSG") <> "", aIntPar.value("COL", "LISTMSG"), "INT-MSG")
        Dim aMsgClmnNr As Integer = 0
        Dim aMatClmnNr As Integer = If(aIntPar.value("COLNR", "DATMAT") <> "", CInt(aIntPar.value("COLNR", "DATMAT")), 1)
        Dim aRetStr As String
        Dim aOKMsg As String = If(aIntPar.value("RET", "OKMSG") <> "", aIntPar.value("RET", "OKMSG"), "OK")

        Dim aWB As Excel.Workbook
        aWB = Globals.SapMmMdExcelAddin.Application.ActiveWorkbook
        Try
            aMLiWs = aWB.Worksheets(aMLiWsName)
        Catch Exc As System.Exception
            MsgBox("No " & aMLiWsName & " Sheet in current workbook. Check if the current workbook is a valid SAP Materal Template",
                   MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAP Materal")
            Exit Sub
        End Try
        Try
            aSLWs = aWB.Worksheets(aSLWsName)
        Catch Exc As System.Exception
            MsgBox("No " & aSLWsName & " Sheet in current workbook. Check if the current workbook is a valid SAP Materal Template",
                   MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAP Materal")
            Exit Sub
        End Try
        parseHeaderLine(aMLiWs, jMax, aMsgClmn, aMsgClmnNr, pHdrLine:=aSLLOff - 3)
        Try
            log.Debug("SapMmMdRibbon_SL.Read - " & "processing data - disabling events, screen update, cursor")
            Globals.SapMmMdExcelAddin.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlWait
            Globals.SapMmMdExcelAddin.Application.EnableEvents = False
            Globals.SapMmMdExcelAddin.Application.ScreenUpdating = False
            Dim i As UInt64 = aMLiLOff + 1
            Dim iOut As ULong = aSLLOff + 1
            Dim jMaxOut As ULong = 0
            Dim aKey As String
            Dim aFirst As Boolean = True
            Dim aClear As Boolean = False
            Dim aFieldArray() As String = {}
            Dim aIsValueArray() As String = {}
            Dim aMLiItem As TDataRec
            Do
                If Left(CStr(aMLiWs.Cells(i, aMsgClmnNr).Value), Len(aOKMsg)) <> aOKMsg Then
                    aKey = CStr(i)
                    Dim aMLiItems As New TData(aIntPar)
                    aMLiItems.ws_parse_line_simple(aMLiWs, aMLiLOff, i, jMax)
                    aMLiItem = aMLiItems.aTDataDic(CStr(i))
                    Dim aTSAP_SourceListData As New TSAP_SourceListData(aPar, aIntPar, aSAPSourceList, "Read")
                    If aTSAP_SourceListData.fillHeader(aMLiItems) Then
                        log.Debug("SapMmMdRibbon_SL.Read - " & "calling aSAPSourceList.Read")
                        aRetStr = aSAPSourceList.Read(aTSAP_SourceListData, pOKMsg:=aOKMsg)
                        log.Debug("SapMmMdRibbon_SL.Read - " & "aSAPSourceList.Read returned, aRetStr=" & aRetStr)
                        aMLiWs.Cells(i, aMsgClmnNr) = CStr(aRetStr)
                        If Left(aRetStr, Len(aOKMsg)) = aOKMsg Then
                            Dim iOutOffset As ULong = 0
                            Dim iOutOffsetMax As ULong = 0
                            Dim aTData As New TData(pPar:=aIntPar)
                            If aFirst Then
                                aClear = True
                                aFirst = False
                                iOut = aSLLOff + 1
                                jMaxOut = aTData.getFieldArray(pWs:=aSLWs, pFieldArray:=aFieldArray, pIsValueArray:=aIsValueArray, pCoff:=0)
                            Else
                                aClear = False
                            End If
                            ' output tables
                            iOutOffset = aTSAP_SourceListData.ws_output(pStructure:="ET_EORD", pFieldArray:=aFieldArray, pIsValueArray:=aIsValueArray, pWs:=aSLWs, pDataKey:="", i:=iOut, jMax:=jMaxOut, pClear:=aClear, pKey:=aMLiItem.getMaterial().Value)
                            iOutOffsetMax = If(iOutOffset > iOutOffsetMax, iOutOffset, iOutOffsetMax)
                            ' iOutOffsetMax = If(iOutOffsetMax = 0, 1, iOutOffsetMax)
                            iOut += iOutOffsetMax
                        End If
                    End If
                End If
                i += 1
            Loop While CStr(aMLiWs.Cells(i, 1).value) <> ""
            log.Debug("SapMmMdRibbon_SL.Read - " & "all data processed - enabling events, screen update, cursor")
            Globals.SapMmMdExcelAddin.Application.EnableEvents = True
            Globals.SapMmMdExcelAddin.Application.ScreenUpdating = True
            Globals.SapMmMdExcelAddin.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault
        Catch ex As System.Exception
            Globals.SapMmMdExcelAddin.Application.EnableEvents = True
            Globals.SapMmMdExcelAddin.Application.ScreenUpdating = True
            Globals.SapMmMdExcelAddin.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault
            MsgBox("SapMmMdRibbon_SL.Read failed! " & ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAP Materal AddIn")
            log.Error("SapMmMdRibbon_SL.Read - " & "Exception=" & ex.ToString)
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
