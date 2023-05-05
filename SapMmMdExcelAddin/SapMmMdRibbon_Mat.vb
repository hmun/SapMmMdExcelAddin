' Copyright 2022 Hermann Mundprecht
' This file is licensed under the terms of the license 'CC BY 4.0'. 
' For a human readable version of the license, see https://creativecommons.org/licenses/by/4.0/

Public Class SapMmMdRibbon_Mat
    Private Shared ReadOnly log As log4net.ILog = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType)
    Public Function getGenParameters(ByRef pPar As SAPCommon.TStr) As Integer
        Dim aPws As Excel.Worksheet
        Dim aWB As Excel.Workbook
        Dim aKey As String
        Dim aName As String
        Dim i As Integer
        log.Debug("SapMmMdRibbon_Mat getGenParametrs - " & "reading Parameter")
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

    Public Sub Change(ByRef pSapCon As SapCon)
        Dim aSAPMaterial As New SAPMaterial(pSapCon)

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
        Dim aMatLOff As Integer = If(aIntPar.value("LOFF", "MAT_DATA") <> "", CInt(aIntPar.value("LOFF", "MAT_DATA")), 4)
        Dim aMatWsName As String = If(aIntPar.value("WS", "MAT_DATA") <> "", aIntPar.value("WS", "MAT_DATA"), "Data")
        Dim aMatWs As Excel.Worksheet
        Dim aMsgClmn As String = If(aIntPar.value("COL", "DATAMSG") <> "", aIntPar.value("COL", "DATAMSG"), "INT-MSG")
        Dim aMsgClmnNr As Integer = 0
        Dim aMatClmnNr As Integer = If(aIntPar.value("COLNR", "DATAMAT") <> "", CInt(aIntPar.value("COLNR", "DATAMAT")), 1)
        Dim aRetStr As String
        Dim aOKMsg As String = If(aIntPar.value("RET", "OKMSG") <> "", aIntPar.value("RET", "OKMSG"), "OK")
        Dim aPriceUseChange As Boolean = If(aIntPar.value("PRICE", "USECHANGE") = "X", True, False)

        Dim aWB As Excel.Workbook
        aWB = Globals.SapMmMdExcelAddin.Application.ActiveWorkbook
        Try
            aMatWs = aWB.Worksheets(aMatWsName)
        Catch Exc As System.Exception
            MsgBox("No " & aMatWsName & " Sheet in current workbook. Check if the current workbook is a valid SAP Materal Template",
                   MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAP Materal")
            Exit Sub
        End Try
        parseHeaderLine(aMatWs, jMax, aMsgClmn, aMsgClmnNr, pHdrLine:=aMatLOff - 3)
        Try
            log.Debug("SapMmMdRibbon_Mat.Change - " & "processing data - disabling events, screen update, cursor")
            Globals.SapMmMdExcelAddin.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlWait
            Globals.SapMmMdExcelAddin.Application.EnableEvents = False
            Globals.SapMmMdExcelAddin.Application.ScreenUpdating = False
            Dim i As UInt64 = aMatLOff + 1
            Dim aKey As String
            Dim aMatItems As New TData(aIntPar)
            Dim aTSAP_MatData As New TSAP_MatData(aPar, aIntPar, aSAPMaterial, "Change")
            Do
                If Left(CStr(aMatWs.Cells(i, aMsgClmnNr).Value), Len(aOKMsg)) <> aOKMsg Then
                    aKey = CStr(i)
                    aRetStr = ""
                    ' read DATA
                    aMatItems.ws_parse_line_simple(aMatWs, aMatLOff, i, jMax)
                    If CStr(aMatWs.Cells(i, aMatClmnNr).Value) <> CStr(aMatWs.Cells(i + 1, aMatClmnNr).Value) Then
                        If aTSAP_MatData.fillHeader(aMatItems) And aTSAP_MatData.fillData(aMatItems) Then
                            If Not aPriceUseChange Then
                                log.Debug("SapMmMdRibbon_Mat.Change - " & "calling aSAPMaterial.Change")
                                aRetStr = aSAPMaterial.Change(aTSAP_MatData, pOKMsg:=aOKMsg)
                                log.Debug("SapMmMdRibbon_Mat.Change - " & "aSAPMaterial.Change returned, aRetStr=" & aRetStr)
                                For Each aKey In aMatItems.aTDataDic.Keys
                                    aMatWs.Cells(CInt(aKey), aMsgClmnNr) = CStr(aRetStr)
                                Next
                            Else
                                log.Debug("SapMmMdRibbon_Mat.Change - " & "using price change logic if valuation view does not exist")
                                Dim aMaterial As SAPCommon.TStrRec = aTSAP_MatData.getMaterialRec()
                                If Not aMaterial Is Nothing Then
                                    Dim aValArea As SAPCommon.TStrRec = aTSAP_MatData.getHdrField("VALUATIONDATA-VAL_AREA")
                                    Dim aValType As SAPCommon.TStrRec = aTSAP_MatData.getHdrField("VALUATIONDATA-VAL_TYPE")
                                    Dim aMaintStat As String = ""
                                    Dim aValAreaRet As String = ""
                                    Dim aValTypeRet As String = ""
                                    Dim aValMaintStatRet As String = ""
                                    Dim aStdPrice As String = ""
                                    Dim aMaintStatRet As String = GetMaintStat(pSapCon:=pSapCon, pMaterial:=aMaterial, aValArea, aValType, aMaintStat, aValAreaRet, aValTypeRet, aValMaintStatRet, aStdPrice)
                                    Dim aStdPr As Double
                                    Try
                                        aStdPr = CDbl(aStdPrice)
                                    Catch ex As Exception
                                        aStdPr = 0
                                    End Try
                                    If Left(aMaintStatRet, Len(aOKMsg)) <> aOKMsg Then
                                        aRetStr = aMaintStatRet
                                    Else
                                        If InStr(aMaintStat, "B") = 0 Or aStdPr = 0 Then  'Accounting view does not exists
                                            log.Debug("SapMmMdRibbon_Mat.Change - " & "accounting view for material " & aMaterial.Value & " does not exist")
                                            Dim aSAPMaterialPrice As New SAPMaterialPrice(pSapCon)
                                            Dim aTSAP_MatPriceData As New TSAP_MatPriceData(aPar, aIntPar, aSAPMaterialPrice, "Change")
                                            If aTSAP_MatPriceData.fromTSAP_MatData(aTSAP_MatData) Then
                                                log.Debug("SapMmMdRibbon_Mat.Change - " & "first calling aSAPMaterial.Change")
                                                aRetStr = aSAPMaterial.Change(aTSAP_MatData, pOKMsg:=aOKMsg)
                                                log.Debug("SapMmMdRibbon_Mat.Change - " & "aSAPMaterial.Change returned, aRetStr=" & aRetStr)
                                                If Left(aRetStr, Len(aOKMsg)) = aOKMsg Then
                                                    log.Debug("SapMmMdRibbon_Mat.Change - " & "then calling aSAPMaterialPrice.Change")
                                                    aRetStr = aRetStr & ";" & aSAPMaterialPrice.Change(aTSAP_MatPriceData, pOKMsg:=aOKMsg)
                                                    log.Debug("SapMmMdRibbon_Mat.Change - " & "aSAPMaterialPrice.Change returned, aRetStr=" & aRetStr)
                                                End If
                                            Else
                                                log.Debug("SapMmMdRibbon_Mat.Change - " & "aTSAP_MatPriceData.fromTSAP_MatData returned False")
                                                aRetStr = "Error: Could not create aTSAP_MatPriceData from TSAP_MatData"
                                            End If
                                        Else
                                            log.Debug("SapMmMdRibbon_Mat.Change - " & "accounting view for material " & aMaterial.Value & "exists")
                                            log.Debug("SapMmMdRibbon_Mat.Change - " & "calling aSAPMaterial.Change")
                                            aRetStr = aSAPMaterial.Change(aTSAP_MatData, pOKMsg:=aOKMsg)
                                            log.Debug("SapMmMdRibbon_Mat.Change - " & "aSAPMaterial.Change returned, aRetStr=" & aRetStr)
                                        End If
                                    End If
                                Else
                                    aRetStr = "Could not find Material-Number in aTSAP_MatData"
                                End If
                                For Each aKey In aMatItems.aTDataDic.Keys
                                    aMatWs.Cells(CInt(aKey), aMsgClmnNr) = CStr(aRetStr)
                                Next
                            End If
                            aMatItems = New TData(aIntPar)
                            aTSAP_MatData = New TSAP_MatData(aPar, aIntPar, aSAPMaterial, "Change")
                        End If
                    End If
                End If
                i += 1
            Loop While Not String.IsNullOrEmpty(CStr(aMatWs.Cells(i, 1).value))
            log.Debug("SapMmMdRibbon_Mat.Change - " & "all data processed - enabling events, screen update, cursor")
            Globals.SapMmMdExcelAddin.Application.EnableEvents = True
            Globals.SapMmMdExcelAddin.Application.ScreenUpdating = True
            Globals.SapMmMdExcelAddin.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault
        Catch ex As System.Exception
            Globals.SapMmMdExcelAddin.Application.EnableEvents = True
            Globals.SapMmMdExcelAddin.Application.ScreenUpdating = True
            Globals.SapMmMdExcelAddin.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault
            MsgBox("SapMmMdRibbon_Mat.Change failed! " & ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAP Materal AddIn")
            log.Error("SapMmMdRibbon_Mat.Change - " & "Exception=" & ex.ToString)
            Exit Sub
        End Try

    End Sub

    Public Sub GetAll(ByRef pSapCon As SapCon)
        Dim aSAPMaterial As New SAPMaterial(pSapCon)

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
        Dim aMLiLOff As Integer = If(aIntPar.value("LOFF", "MAT_LIST") <> "", CInt(aIntPar.value("LOFF", "MAT_LIST")), 4)
        Dim aMatLOff As Integer = If(aIntPar.value("LOFF", "MAT_DATA") <> "", CInt(aIntPar.value("LOFF", "MAT_DATA")), 4)
        Dim aMLiWsName As String = If(aIntPar.value("WS", "MAT_LIST") <> "", aIntPar.value("WS", "MAT_LIST"), "Material_List")
        Dim aMatWsName As String = If(aIntPar.value("WS", "MAT_DATA") <> "", aIntPar.value("WS", "MAT_DATA"), "Data")
        Dim aMatWs As Excel.Worksheet
        Dim aMLiWs As Excel.Worksheet
        Dim aMsgClmn As String = If(aIntPar.value("COL", "LISTMSG") <> "", aIntPar.value("COL", "LISTMSG"), "INT-MSG")
        Dim aMsgClmnNr As Integer = 0
        Dim aMatClmnNr As Integer = If(aIntPar.value("COLNR", "DATAMAT") <> "", CInt(aIntPar.value("COLNR", "DATAMAT")), 1)
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
            aMatWs = aWB.Worksheets(aMatWsName)
        Catch Exc As System.Exception
            MsgBox("No " & aMatWsName & " Sheet in current workbook. Check if the current workbook is a valid SAP Materal Template",
                   MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAP Materal")
            Exit Sub
        End Try
        parseHeaderLine(aMLiWs, jMax, aMsgClmn, aMsgClmnNr, pHdrLine:=aMatLOff - 3)
        Try
            log.Debug("SapMmMdRibbon_Mat.GetAll - " & "processing data - disabling events, screen update, cursor")
            Globals.SapMmMdExcelAddin.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlWait
            Globals.SapMmMdExcelAddin.Application.EnableEvents = False
            Globals.SapMmMdExcelAddin.Application.ScreenUpdating = False
            Dim i As UInt64 = aMLiLOff + 1
            Dim iOut As ULong = aMatLOff + 1
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
                    Dim aTSAP_MatData As New TSAP_MatData(aPar, aIntPar, aSAPMaterial, "GetAll")
                    If aTSAP_MatData.fillHeader(aMLiItems) Then
                        log.Debug("SapMmMdRibbon_Mat.GetAll - " & "calling aSAPMaterial.GetAll")
                        aRetStr = aSAPMaterial.GetAll(aTSAP_MatData, pOKMsg:=aOKMsg)
                        log.Debug("SapMmMdRibbon_Mat.GetAll - " & "aSAPMaterial.GetAll returned, aRetStr=" & aRetStr)
                        aMLiWs.Cells(i, aMsgClmnNr) = CStr(aRetStr)
                        If Left(aRetStr, Len(aOKMsg)) = aOKMsg Then
                            Dim iOutOffset As ULong = 0
                            Dim iOutOffsetMax As ULong = 0
                            Dim aTData As New TData(pPar:=aIntPar)
                            If aFirst Then
                                aClear = True
                                aFirst = False
                                iOut = aMatLOff + 1
                                jMaxOut = aTData.getFieldArray(pWs:=aMatWs, pFieldArray:=aFieldArray, pIsValueArray:=aIsValueArray, pCoff:=0)
                            Else
                                aClear = False
                            End If
                            ' output structures
                            aTSAP_MatData.ws_output_line(pStructure:="CLIENTDATA", pFieldArray:=aFieldArray, pIsValueArray:=aIsValueArray, pWs:=aMatWs, pDataKey:="", i:=iOut, jMax:=jMaxOut, pClear:=aClear, pKey:="")
                            aTSAP_MatData.ws_output_line(pStructure:="PLANTDATA", pFieldArray:=aFieldArray, pIsValueArray:=aIsValueArray, pWs:=aMatWs, pDataKey:="", i:=iOut, jMax:=jMaxOut, pClear:=aClear, pKey:="")
                            aTSAP_MatData.ws_output_line(pStructure:="FORECASTPARAMETERS", pFieldArray:=aFieldArray, pIsValueArray:=aIsValueArray, pWs:=aMatWs, pDataKey:="", i:=iOut, jMax:=jMaxOut, pClear:=aClear, pKey:="")
                            aTSAP_MatData.ws_output_line(pStructure:="PLANNINGDATA", pFieldArray:=aFieldArray, pIsValueArray:=aIsValueArray, pWs:=aMatWs, pDataKey:="", i:=iOut, jMax:=jMaxOut, pClear:=aClear, pKey:="")
                            aTSAP_MatData.ws_output_line(pStructure:="VALUATIONDATA", pFieldArray:=aFieldArray, pIsValueArray:=aIsValueArray, pWs:=aMatWs, pDataKey:="", i:=iOut, jMax:=jMaxOut, pClear:=aClear, pKey:="")
                            aTSAP_MatData.ws_output_line(pStructure:="WAREHOUSENUMBERDATA", pFieldArray:=aFieldArray, pIsValueArray:=aIsValueArray, pWs:=aMatWs, pDataKey:="", i:=iOut, jMax:=jMaxOut, pClear:=aClear, pKey:="")
                            aTSAP_MatData.ws_output_line(pStructure:="SALESDATA", pFieldArray:=aFieldArray, pIsValueArray:=aIsValueArray, pWs:=aMatWs, pDataKey:="", i:=iOut, jMax:=jMaxOut, pClear:=aClear, pKey:="")
                            aTSAP_MatData.ws_output_line(pStructure:="STORAGETYPEDATA", pFieldArray:=aFieldArray, pIsValueArray:=aIsValueArray, pWs:=aMatWs, pDataKey:="", i:=iOut, jMax:=jMaxOut, pClear:=aClear, pKey:="")
                            aTSAP_MatData.ws_output_line(pStructure:="PRTDATA", pFieldArray:=aFieldArray, pIsValueArray:=aIsValueArray, pWs:=aMatWs, pDataKey:="", i:=iOut, jMax:=jMaxOut, pClear:=aClear, pKey:="")
                            aTSAP_MatData.ws_output_line(pStructure:="LIFOVALUATIONDATA", pFieldArray:=aFieldArray, pIsValueArray:=aIsValueArray, pWs:=aMatWs, pDataKey:="", i:=iOut, jMax:=jMaxOut, pClear:=aClear, pKey:="")
                            ' output tables
                            iOutOffset = aTSAP_MatData.ws_output(pStructure:="MATERIALDESCRIPTION", pFieldArray:=aFieldArray, pIsValueArray:=aIsValueArray, pWs:=aMatWs, pDataKey:="", i:=iOut, jMax:=jMaxOut, pClear:=aClear, pKey:=aMLiItem.getMaterial().Value)
                            iOutOffsetMax = If(iOutOffset > iOutOffsetMax, iOutOffset, iOutOffsetMax)
                            iOutOffset = aTSAP_MatData.ws_output(pStructure:="UNITSOFMEASURE", pFieldArray:=aFieldArray, pIsValueArray:=aIsValueArray, pWs:=aMatWs, pDataKey:="", i:=iOut, jMax:=jMaxOut, pClear:=aClear, pKey:=aMLiItem.getMaterial().Value)
                            iOutOffsetMax = If(iOutOffset > iOutOffsetMax, iOutOffset, iOutOffsetMax)
                            iOutOffset = aTSAP_MatData.ws_output(pStructure:="INTERNATIONALARTNOS", pFieldArray:=aFieldArray, pIsValueArray:=aIsValueArray, pWs:=aMatWs, pDataKey:="", i:=iOut, jMax:=jMaxOut, pClear:=aClear, pKey:=aMLiItem.getMaterial().Value)
                            iOutOffsetMax = If(iOutOffset > iOutOffsetMax, iOutOffset, iOutOffsetMax)
                            iOutOffset = aTSAP_MatData.ws_output(pStructure:="TAXCLASSIFICATIONS", pFieldArray:=aFieldArray, pIsValueArray:=aIsValueArray, pWs:=aMatWs, pDataKey:="", i:=iOut, jMax:=jMaxOut, pClear:=aClear, pKey:=aMLiItem.getMaterial().Value)
                            iOutOffsetMax = If(iOutOffset > iOutOffsetMax, iOutOffset, iOutOffsetMax)
                            iOutOffset = aTSAP_MatData.ws_output(pStructure:="EXTENSIONOUT", pFieldArray:=aFieldArray, pIsValueArray:=aIsValueArray, pWs:=aMatWs, pDataKey:="", i:=iOut, jMax:=jMaxOut, pClear:=aClear, pKey:=aMLiItem.getMaterial().Value)
                            iOutOffsetMax = If(iOutOffset > iOutOffsetMax, iOutOffset, iOutOffsetMax)
                            iOutOffsetMax = If(iOutOffsetMax = 0, 1, iOutOffsetMax)
                            iOut += iOutOffsetMax
                        End If
                    End If
                End If
                i += 1
            Loop While CStr(aMLiWs.Cells(i, 1).value) <> ""
            log.Debug("SapMmMdRibbon_Mat.GetAll - " & "all data processed - enabling events, screen update, cursor")
            Globals.SapMmMdExcelAddin.Application.EnableEvents = True
            Globals.SapMmMdExcelAddin.Application.ScreenUpdating = True
            Globals.SapMmMdExcelAddin.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault
        Catch ex As System.Exception
            Globals.SapMmMdExcelAddin.Application.EnableEvents = True
            Globals.SapMmMdExcelAddin.Application.ScreenUpdating = True
            Globals.SapMmMdExcelAddin.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault
            MsgBox("SapMmMdRibbon_Mat.GetAll failed! " & ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAP Materal AddIn")
            log.Error("SapMmMdRibbon_Mat.GetAll - " & "Exception=" & ex.ToString)
            Exit Sub
        End Try
    End Sub

    Public Function GetMaintStat(ByRef pSapCon As SapCon, pMaterial As SAPCommon.TStrRec, pValArea As SAPCommon.TStrRec, pValType As SAPCommon.TStrRec,
                                 ByRef pMaintStat As String, ByRef pValAreaRet As String, ByRef pValTypeRet As String, ByRef pValMaintStatRet As String, ByRef pStdPrice As String,
                                 Optional pFieldName As String = "") As String
        Dim aSAPMaterial As New SAPMaterial(pSapCon)

        Dim aPar As New SAPCommon.TStr
        Dim aIntPar As New SAPCommon.TStr

        GetMaintStat = ""
        ' get general parameters
        If getGenParameters(aPar) = False Then
            Exit Function
        End If
        ' get internal parameters
        If Not getIntParameters(aIntPar) Then
            Exit Function
        End If
        Dim aMatFieldname As String
        If String.IsNullOrEmpty(pFieldName) Then
            aMatFieldname = If(aIntPar.value("MAT", "FIELD_GET") <> "", aIntPar.value("MAT", "FIELD_GET"), "MATERIAL")
        Else
            aMatFieldname = pFieldName
        End If
        If InStr(aMatFieldname, "-") = 0 Then
            aMatFieldname = "-" & aMatFieldname
        End If
        Dim aRetStr As String
        Dim aOKMsg As String = If(aIntPar.value("RET", "OKMSG") <> "", aIntPar.value("RET", "OKMSG"), "OK")
        Try
            log.Debug("SapMmMdRibbon_Mat.GetMaintStat - " & "processing data")
            Dim aKey As String
            Dim aFirst As Boolean = True
            Dim aFieldArray() As String = {}
            Dim aIsValueArray() As String = {}
            aKey = "0"
            Dim aMLiItems As New TData(aIntPar)
            aMLiItems.addValue(aKey, aMatFieldname, pMaterial.Value, "", pMaterial.Format, pEmptyChar:="", pEmty:=False)
            If Not pValArea Is Nothing Then
                aMLiItems.addValue(aKey, "-VAL_AREA", pValArea.Value, "", pValArea.Format, pEmptyChar:="", pEmty:=False)
            End If
            If Not pValType Is Nothing Then
                aMLiItems.addValue(aKey, "-VAL_TYPE", pValType.Value, "", pValType.Format, pEmptyChar:="", pEmty:=False)
            End If
            Dim aTSAP_MatData As New TSAP_MatData(aPar, aIntPar, aSAPMaterial, "GetAll")
            If aTSAP_MatData.fillHeader(aMLiItems) Then
                log.Debug("SapMmMdRibbon_Mat.GetMaintStat - " & "calling aSAPMaterial.GetAll")
                aRetStr = aSAPMaterial.GetAll(aTSAP_MatData, pOKMsg:=aOKMsg)
                log.Debug("SapMmMdRibbon_Mat.GetMaintStat - " & "aSAPMaterial.GetAll returned, aRetStr=" & aRetStr)
                pMaintStat = aTSAP_MatData.getDataField(pStrucName:="CLIENTDATA", pFieldname:="MAINT_STAT").Value
                pValAreaRet = aTSAP_MatData.getDataField(pStrucName:="VALUATIONDATA", pFieldname:="VAL_AREA").Value
                pValTypeRet = aTSAP_MatData.getDataField(pStrucName:="VALUATIONDATA", pFieldname:="VAL_TYPE").Value
                pStdPrice = aTSAP_MatData.getDataField(pStrucName:="VALUATIONDATA", pFieldname:="STD_PRICE").Value
                pValMaintStatRet = aTSAP_MatData.getDataField(pStrucName:="VALUATIONDATA", pFieldname:="MAINT_STAT").Value
                GetMaintStat = aRetStr
            End If
            log.Debug("SapMmMdRibbon_Mat.GetMaintStat - " & "all data processed")
        Catch ex As System.Exception
            MsgBox("SapMmMdRibbon_Mat.GetMaintStat failed! " & ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAP Materal AddIn")
            log.Error("SapMmMdRibbon_Mat.GetMaintStat - " & "Exception=" & ex.ToString)
            GetMaintStat = ""
            Exit Function
        End Try
    End Function

    Public Sub PriceChange(ByRef pSapCon As SapCon)
        Dim aSAPMaterialPrice As New SAPMaterialPrice(pSapCon)

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
        Dim aMatLOff As Integer = If(aIntPar.value("LOFF", "MAT_DATA") <> "", CInt(aIntPar.value("LOFF", "MAT_DATA")), 4)
        Dim aMatWsName As String = If(aIntPar.value("WS", "MAT_DATA") <> "", aIntPar.value("WS", "MAT_DATA"), "Data")
        Dim aMatWs As Excel.Worksheet
        Dim aMsgClmn As String = If(aIntPar.value("COL", "DATAMSG") <> "", aIntPar.value("COL", "DATAMSG"), "INT-MSG")
        Dim aMsgClmnNr As Integer = 0
        Dim aMatClmnNr As Integer = If(aIntPar.value("COLNR", "DATAMAT") <> "", CInt(aIntPar.value("COLNR", "DATAMAT")), 1)
        Dim aRetStr As String
        Dim aOKMsg As String = If(aIntPar.value("RET", "OKMSG") <> "", aIntPar.value("RET", "OKMSG"), "OK")

        Dim aWB As Excel.Workbook
        aWB = Globals.SapMmMdExcelAddin.Application.ActiveWorkbook
        Try
            aMatWs = aWB.Worksheets(aMatWsName)
        Catch Exc As System.Exception
            MsgBox("No " & aMatWsName & " Sheet in current workbook. Check if the current workbook is a valid SAP Materal Template",
                   MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAP Materal")
            Exit Sub
        End Try
        parseHeaderLine(aMatWs, jMax, aMsgClmn, aMsgClmnNr, pHdrLine:=aMatLOff - 3)
        Try
            log.Debug("SapMmMdRibbon_Mat.PriceChange - " & "processing data - disabling events, screen update, cursor")
            Globals.SapMmMdExcelAddin.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlWait
            Globals.SapMmMdExcelAddin.Application.EnableEvents = False
            Globals.SapMmMdExcelAddin.Application.ScreenUpdating = False
            Dim i As UInt64 = aMatLOff + 1
            Dim aKey As String
            Dim aMatItems As New TData(aIntPar)
            Dim aTSAP_MatPriceData As New TSAP_MatPriceData(aPar, aIntPar, aSAPMaterialPrice, "Change")
            Do
                If Left(CStr(aMatWs.Cells(i, aMsgClmnNr).Value), Len(aOKMsg)) <> aOKMsg Then
                    aKey = CStr(i)
                    aRetStr = ""
                    ' read DATA
                    aMatItems.ws_parse_line_simple(aMatWs, aMatLOff, i, jMax)
                    If CStr(aMatWs.Cells(i, aMatClmnNr).Value) <> CStr(aMatWs.Cells(i + 1, aMatClmnNr).Value) Then
                        If aTSAP_MatPriceData.fillHeader(aMatItems) And aTSAP_MatPriceData.fillData(aMatItems) Then
                            log.Debug("SapMmMdRibbon_Mat.PriceChange - " & "calling aSAPMaterialPrice.Change")
                            aRetStr = aSAPMaterialPrice.Change(aTSAP_MatPriceData, pOKMsg:=aOKMsg)
                            log.Debug("SapMmMdRibbon_Mat.PriceChange - " & "aSAPMaterialPrice.Change returned, aRetStr=" & aRetStr)
                            For Each aKey In aMatItems.aTDataDic.Keys
                                aMatWs.Cells(CInt(aKey), aMsgClmnNr) = CStr(aRetStr)
                            Next
                        End If
                    End If
                End If
                i += 1
            Loop While Not String.IsNullOrEmpty(CStr(aMatWs.Cells(i, 1).value))
            log.Debug("SapMmMdRibbon_Mat.PriceChange - " & "all data processed - enabling events, screen update, cursor")
            Globals.SapMmMdExcelAddin.Application.EnableEvents = True
            Globals.SapMmMdExcelAddin.Application.ScreenUpdating = True
            Globals.SapMmMdExcelAddin.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault
        Catch ex As System.Exception
            Globals.SapMmMdExcelAddin.Application.EnableEvents = True
            Globals.SapMmMdExcelAddin.Application.ScreenUpdating = True
            Globals.SapMmMdExcelAddin.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault
            MsgBox("SapMmMdRibbon_Mat.PriceChange failed! " & ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAP Materal AddIn")
            log.Error("SapMmMdRibbon_Mat.PriceChange - " & "Exception=" & ex.ToString)
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
