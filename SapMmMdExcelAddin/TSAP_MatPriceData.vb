' Copyright 2022 Hermann Mundprecht
' This file is licensed under the terms of the license 'CC BY 4.0'. 
' For a human readable version of the license, see https://creativecommons.org/licenses/by/4.0/

Imports SAP.Middleware.Connector

Public Class TSAP_MatPriceData

    Public aHdrRec As TDataRec
    Public aDataDic As TDataDic

    Public aStrucDic As Dictionary(Of String, RfcStructureMetadata)
    Public aParamDic As Dictionary(Of String, RfcParameterMetadata)

    Private aPar As SAPCommon.TStr
    Private aIntPar As SAPCommon.TStr
    Private aFunction As String
    Private Shared ReadOnly log As log4net.ILog = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType)

    Public Sub New(ByRef pPar As SAPCommon.TStr, ByRef pIntPar As SAPCommon.TStr, ByRef pSAPMaterialPrice As SAPMaterialPrice, pFunction As String)
        aPar = pPar
        aIntPar = pIntPar
        aFunction = pFunction
        aDataDic = New TDataDic(aIntPar)
        aHdrRec = New TDataRec(aIntPar)
        aStrucDic = New Dictionary(Of String, RfcStructureMetadata)
        aParamDic = New Dictionary(Of String, RfcParameterMetadata)
        ' get Metadata
        If pFunction = "Change" Then
            pSAPMaterialPrice.getMeta_Change(aParamDic, aStrucDic)
        End If
    End Sub

    Public Function fillHeader(pData As TData) As Boolean
        Dim aKvb As KeyValuePair(Of String, SAPCommon.TStrRec)
        Dim aParDic As Dictionary(Of String, SAPCommon.TStrRec) = aPar.getData()
        Dim aPostRec As TDataRec = pData.getFirstRecord()
        Dim aTStrRec As SAPCommon.TStrRec
        Dim aNewTStrRec As SAPCommon.TStrRec
        Dim aNewHdrRec As New TDataRec(aIntPar)
        Dim aStrucName() As String
        Dim aLen As Integer = 0
        For Each aKvb In aParDic
            aTStrRec = aKvb.Value
            If valid_Import_Field(aTStrRec) Then
                aNewHdrRec.setValues(aTStrRec.getKey(), aTStrRec.Value, aTStrRec.Currency, aTStrRec.Format, pEmptyChar:="")
            End If
            aStrucName = Split(aTStrRec.Strucname, "+")
            For s As Integer = 0 To aStrucName.Length - 1
                aNewTStrRec = New SAPCommon.TStrRec(aStrucName(s), aTStrRec.Fieldname, aTStrRec.Value, aTStrRec.Currency, aTStrRec.Format)
                If valid_Structure_Field(aNewTStrRec) Then
                    aNewHdrRec.setValues(aNewTStrRec.getKey(), aNewTStrRec.Value, aNewTStrRec.Currency, aNewTStrRec.Format, pEmptyChar:="")
                End If
            Next
        Next
        ' First fill the value from the paramters and tehn overwrite them from the posting record
        For Each aTStrRec In aPostRec.aTDataRecCol
            If valid_Import_Field(aTStrRec) Then
                aNewHdrRec.setValues(aTStrRec.getKey(), aTStrRec.Value, aTStrRec.Currency, aTStrRec.Format)
            End If
            aStrucName = Split(aTStrRec.Strucname, "+")
            For s As Integer = 0 To aStrucName.Length - 1
                aNewTStrRec = New SAPCommon.TStrRec(aStrucName(s), aTStrRec.Fieldname, aTStrRec.Value, aTStrRec.Currency, aTStrRec.Format)
                If valid_Structure_Field(aNewTStrRec) Then
                    aNewHdrRec.setValues(aNewTStrRec.getKey(), aNewTStrRec.Value, aNewTStrRec.Currency, aNewTStrRec.Format, pEmptyChar:="")
                End If
            Next
        Next
        aHdrRec = aNewHdrRec
        fillHeader = True
    End Function

    Public Function fillData(pData As TData) As Boolean
        Dim aParDic As Dictionary(Of String, SAPCommon.TStrRec) = aPar.getData()
        Dim aParKvb As KeyValuePair(Of String, SAPCommon.TStrRec)
        Dim aKvB As KeyValuePair(Of String, TDataRec)
        Dim aTDataRec As TDataRec
        Dim aTStrRec As SAPCommon.TStrRec
        Dim aCurTStrRec As SAPCommon.TStrRec
        Dim aCnt As UInt64
        Dim aPostRec As TDataRec = pData.getFirstRecord()
        Dim aStdPriceRec As SAPCommon.TStrRec = If(Not aPostRec Is Nothing, aPostRec.getStdPrice(), Nothing)
        Dim aPriceUnitRec As SAPCommon.TStrRec = If(Not aPostRec Is Nothing, aPostRec.getPriceUnit(), Nothing)
        Dim aUseStdPrice As Boolean = If(Not aStdPriceRec Is Nothing, True, False)
        Dim aUseCurPrice As Boolean = If(Not aPostRec Is Nothing, aPostRec.hasCurPrice(), False)
        aDataDic = New TDataDic(aIntPar)
        fillData = False
        If aUseStdPrice Then
            Dim aFxRate As Double
            Dim aPrice As Double
            Dim aPriceUnit As Integer = If(Not aPriceUnitRec Is Nothing, CInt(aPriceUnitRec.Value), 1)
            ' data uses Standard Price and currency keys, fx rates from parameters
            Dim aPriceTStrRec = aPostRec.getStdPrice()
            aPrice = If(Not aPriceTStrRec Is Nothing, CDbl(aPriceTStrRec.Value), 0)
            aCnt = 0
            For Each aParKvb In aParDic
                aTStrRec = aParKvb.Value
                If valid_Cur_Field(aTStrRec) Then
                    aCnt += 1
                    ' fill a price record
                    If aParDic.Keys.Contains(aTStrRec.Strucname & "-" & "VALUATION_VIEW") Then
                        aCurTStrRec = aParDic(aTStrRec.Strucname & "-" & "VALUATION_VIEW")
                        aDataDic.addValue(CStr(aCnt), New SAPCommon.TStrRec("PRICES", "VALUATION_VIEW", aCurTStrRec.Value, aCurTStrRec.Currency, aCurTStrRec.Format))
                    Else
                        aDataDic.addValue(CStr(aCnt), New SAPCommon.TStrRec("PRICES", "VALUATION_VIEW", "0", "", ""))
                    End If
                    aDataDic.addValue(CStr(aCnt), New SAPCommon.TStrRec("PRICES", "CURR_TYPE", Right(aTStrRec.Strucname, 2), "", ""))
                    aDataDic.addValue(CStr(aCnt), New SAPCommon.TStrRec("PRICES", "CURRENCY", aTStrRec.Value, "", ""))
                    aFxRate = If(aParDic.Keys.Contains(aTStrRec.Strucname & "-" & "FXRATE"), CDbl(aParDic(aTStrRec.Strucname & "-" & "FXRATE").Value), 1)
                    aDataDic.addValue(CStr(aCnt), New SAPCommon.TStrRec("PRICES", "PRICE", CStr(Math.Round(aPrice * aFxRate, 2)), "X", ""))
                    aDataDic.addValue(CStr(aCnt), New SAPCommon.TStrRec("PRICES", "PRICE_UNIT", aPriceUnit, "X", ""))
                End If
            Next
        ElseIf aUseCurPrice Then
            ' data uses prices in culumns and currency keys from parameter
            Dim aPriceUnit As Integer = If(Not aPriceUnitRec Is Nothing, CInt(aPriceUnitRec.Value), 1)
            aCnt = 0
            For Each aTStrRec In aPostRec.aTDataRecCol
                If valid_Price_Field(aTStrRec) Then
                    aCnt += 1
                    ' fill a price record
                    If aParDic.Keys.Contains(aTStrRec.Strucname & "-" & "VALUATION_VIEW") Then
                        aCurTStrRec = aParDic(aTStrRec.Strucname & "-" & "VALUATION_VIEW")
                        aDataDic.addValue(CStr(aCnt), New SAPCommon.TStrRec("PRICES", "VALUATION_VIEW", aCurTStrRec.Value, aCurTStrRec.Currency, aCurTStrRec.Format))
                    Else
                        aDataDic.addValue(CStr(aCnt), New SAPCommon.TStrRec("PRICES", "VALUATION_VIEW", "0", "", ""))
                    End If
                    aDataDic.addValue(CStr(aCnt), New SAPCommon.TStrRec("PRICES", "CURR_TYPE", Right(aTStrRec.Strucname, 2), "", ""))
                    If aParDic.Keys.Contains(aTStrRec.Strucname & "-" & "CURRENCY") Then
                        aCurTStrRec = aParDic(aTStrRec.Strucname & "-" & "CURRENCY")
                        aDataDic.addValue(CStr(aCnt), New SAPCommon.TStrRec("PRICES", "CURRENCY", aCurTStrRec.Value, aCurTStrRec.Currency, aCurTStrRec.Format))
                    Else
                        log.Warn("TSAP_MatPriceData.fillData - " & "no currency key in paramters for " & aTStrRec.Strucname & " exit false")
                        Exit Function
                    End If
                    aDataDic.addValue(CStr(aCnt), New SAPCommon.TStrRec("PRICES", "PRICE", aTStrRec.Value, aTStrRec.Currency, aTStrRec.Format))
                    aDataDic.addValue(CStr(aCnt), New SAPCommon.TStrRec("PRICES", "PRICE_UNIT", aPriceUnit, "X", ""))
                End If
            Next
        Else
            ' regular logic, prices are in lines
            aCnt = 1
            For Each aKvB In pData.aTDataDic
                aTDataRec = aKvB.Value
                For Each aTStrRec In aTDataRec.aTDataRecCol
                    If valid_Table_Field(aTStrRec) Then
                        aDataDic.addValue(CStr(aCnt), aTStrRec)
                    End If
                Next
                aCnt += 1
            Next
        End If
        fillData = True
    End Function

    Public Function fromTSAP_MatData(ByRef aTSAP_MatData As TSAP_MatData) As Boolean
        Dim aPriceTStrRec As SAPCommon.TStrRec
        Dim aPriceUnitTStrRec As SAPCommon.TStrRec
        Dim aCurTStrRec As SAPCommon.TStrRec
        Dim aTStrRec As SAPCommon.TStrRec
        Dim aKvb As KeyValuePair(Of String, SAPCommon.TStrRec)
        Dim aParDic As Dictionary(Of String, SAPCommon.TStrRec)
        Dim aCnt As UInt64
        Dim aFxRate As Double
        Dim aPrice As Double
        Dim aPriceUnit As Integer

        Dim aMatFieldname = If(aIntPar.value("MAT", "FIELD_PRICE") <> "", aIntPar.value("MAT", "FIELD_PRICE"), "MATERIAL")
        If InStr(aMatFieldname, "-") = 0 Then
            aMatFieldname = "-" & aMatFieldname
        End If
        fromTSAP_MatData = False
        ' get the new standard price
        aPriceTStrRec = aTSAP_MatData.getHdrField("VALUATIONDATA-STD_PRICE")
        If aPriceTStrRec Is Nothing Then
            log.Warn("TSAP_MatPriceData.fromTSAP_MatData - " & "no Standard Price found in aTSAP_MatData - exit false")
            Exit Function
        End If
        aPrice = If(Not aPriceTStrRec Is Nothing, CDbl(aPriceTStrRec.Value), 0)
        aPriceUnitTStrRec = aTSAP_MatData.getHdrField("VALUATIONDATA-PRICE_UNIT")
        aPriceUnit = If(aPriceUnitTStrRec Is Nothing, "1", CInt(aPriceUnitTStrRec.Value))
        ' process parameters first
        Dim aCurFound As Boolean = False
        aCnt = 0
        aParDic = aPar.getData()
        For Each aKvb In aParDic
            aTStrRec = aKvb.Value
            If valid_Import_Field(aTStrRec) Or valid_Structure_Field(aTStrRec) Then
                aHdrRec.setValues(aTStrRec.getKey(), aTStrRec.Value, aTStrRec.Currency, aTStrRec.Format, pEmptyChar:="")
            ElseIf valid_Cur_Field(aTStrRec) Then
                aCnt += 1
                ' fill a price record
                If aParDic.Keys.Contains(aTStrRec.Strucname & "-" & "VALUATION_VIEW") Then
                    aCurTStrRec = aParDic(aTStrRec.Strucname & "-" & "VALUATION_VIEW")
                    aDataDic.addValue(CStr(aCnt), New SAPCommon.TStrRec("PRICES", "VALUATION_VIEW", aCurTStrRec.Value, aCurTStrRec.Currency, aCurTStrRec.Format))
                Else
                    aDataDic.addValue(CStr(aCnt), New SAPCommon.TStrRec("PRICES", "VALUATION_VIEW", "0", "", ""))
                End If
                aCurTStrRec = New SAPCommon.TStrRec("PRICES", "CURR_TYPE", Right(aTStrRec.Strucname, 2), "", "")
                aDataDic.addValue(CStr(aCnt), aCurTStrRec)
                aCurTStrRec = New SAPCommon.TStrRec("PRICES", "CURRENCY", aTStrRec.Value, "", "")
                aDataDic.addValue(CStr(aCnt), aCurTStrRec)
                aFxRate = If(aParDic.Keys.Contains(aTStrRec.Strucname & "-" & "FXRATE"), CDbl(aParDic(aTStrRec.Strucname & "-" & "FXRATE").Value), 1)
                aDataDic.addValue(CStr(aCnt), New SAPCommon.TStrRec("PRICES", "PRICE", CStr(Math.Round(aPrice * aFxRate, 2)), "X", ""))
                aDataDic.addValue(CStr(aCnt), New SAPCommon.TStrRec("PRICES", "PRICE_UNIT", CStr(aPriceUnit), "X", ""))
                aCurFound = True
            End If
        Next
        If Not aCurFound Then
            log.Warn("TSAP_MatPriceData.fromTSAP_MatData - " & "no Currency Type definition found in Paramters - exit false")
            Exit Function
        End If
        ' process values from TSAP_MatData
        Dim aMaterialRec As SAPCommon.TStrRec = aTSAP_MatData.getMaterialRec
        aHdrRec.setValues(aMatFieldname, aMaterialRec.Value, aMaterialRec.Currency, aMaterialRec.Format, pEmptyChar:="")
        aTStrRec = aTSAP_MatData.getHdrField("VALUATIONDATA-VAL_AREA")
        If Not aTStrRec Is Nothing Then
            aHdrRec.setValues("-" & "VALUATIONAREA", aTStrRec.Value, aTStrRec.Currency, aTStrRec.Format, pEmptyChar:="")
        Else
            log.Warn("TSAP_MatData.fromTSAP_MatData - " & "no VALUATIONDATA-VAL_AREA found in aTSAP_MatData - exit false")
            Exit Function
        End If
        aTStrRec = aTSAP_MatData.getHdrField("VALUATIONDATA-VAL_TYPE")
        If Not aTStrRec Is Nothing Then
            aHdrRec.setValues("-" & "VALUATIONTYPE", aTStrRec.Value, aTStrRec.Currency, aTStrRec.Format, pEmptyChar:="")
        Else
            aHdrRec.setValues("-" & "VALUATIONTYPE", "", "", "", pEmptyChar:="")
        End If
        fromTSAP_MatData = True
    End Function

    Public Function valid_Import_Field(pTStrRec As SAPCommon.TStrRec) As Boolean
        valid_Import_Field = False
        If pTStrRec.Strucname = "" Or pTStrRec.Strucname = "I" Then
            If aParamDic.ContainsKey("I|" & pTStrRec.Fieldname) Then
                valid_Import_Field = True
            End If
        End If
    End Function

    Public Function valid_Structure_Field(pTStrRec As SAPCommon.TStrRec) As Boolean
        Dim aStrucName() As String
        valid_Structure_Field = False
        aStrucName = Split(pTStrRec.Strucname, "+")
        For s As Integer = 0 To aStrucName.Length - 1
            If aStrucDic.ContainsKey("S|" & aStrucName(s)) Then
                valid_Structure_Field = isInStructure(pTStrRec.Fieldname, aStrucDic("S|" & aStrucName(s)))
            End If
        Next
    End Function

    Public Function valid_Table_Field(pTStrRec As SAPCommon.TStrRec) As Boolean
        Dim aStrucName() As String
        valid_Table_Field = False
        aStrucName = Split(pTStrRec.Strucname, "+")
        For s As Integer = 0 To aStrucName.Length - 1
            If aStrucDic.ContainsKey("T|" & aStrucName(s)) Then
                valid_Table_Field = isInStructure(pTStrRec.Fieldname, aStrucDic("T|" & aStrucName(s)))
            End If
        Next
    End Function
    Public Function valid_Cur_Field(pTStrRec As SAPCommon.TStrRec) As Boolean
        Dim aStrucName() As String
        valid_Cur_Field = False
        aStrucName = Split(pTStrRec.Strucname, "+")
        If isInArray("A00", aStrucName) Or isInArray("A10", aStrucName) Or isInArray("A20", aStrucName) Or isInArray("A30", aStrucName) Or isInArray("A40", aStrucName) Then
            valid_Cur_Field = If(pTStrRec.Fieldname = "CURRENCY", True, False)
        End If
    End Function

    Public Function valid_Price_Field(pTStrRec As SAPCommon.TStrRec) As Boolean
        Dim aStrucName() As String
        valid_Price_Field = False
        aStrucName = Split(pTStrRec.Strucname, "+")
        If isInArray("A00", aStrucName) Or isInArray("A10", aStrucName) Or isInArray("A20", aStrucName) Or isInArray("A30", aStrucName) Or isInArray("A40", aStrucName) Then
            valid_Price_Field = If(pTStrRec.Fieldname = "PRICE", True, False)
        End If
    End Function

    Private Function isInStructure(pName As String, pRfcStructureMetadata As RfcStructureMetadata, Optional ByRef pLen As Integer = 0) As Boolean
        Dim aRfcFieldMetadata As RfcFieldMetadata
        Try
            aRfcFieldMetadata = pRfcStructureMetadata.Item(pName)
            isInStructure = True
            pLen = aRfcFieldMetadata.NucLength
        Catch ex As Exception
            isInStructure = False
            pLen = 0
        End Try
    End Function

    Private Function isInArray(pString As String, pArray As Object) As Boolean
        Dim st As String, M As String
        M = "$"
        st = M & Join(pArray, M) & M
        isInArray = InStr(st, M & pString & M) > 0
        ' isInArray = (UBound(Filter(pArray, pString)) > -1)
    End Function

    Public Sub dumpHeader()
        Dim dumpHd As String = If(aIntPar.value("DBG", "DUMPHEADER") <> "", aIntPar.value("DBG", "DUMPHEADER"), "")
        If dumpHd <> "" Then
            Dim aDWS As Excel.Worksheet
            Dim aWB As Excel.Workbook
            Dim aRange As Excel.Range
            aWB = Globals.SapMmMdExcelAddin.Application.ActiveWorkbook
            Try
                aDWS = aWB.Worksheets(dumpHd)
                aDWS.Activate()
            Catch Exc As System.Exception
                log.Warn("dumpHeader - " & "No " & dumpHd & " Sheet in current workbook.")
                MsgBox("No " & dumpHd & " Sheet in current workbook. Check the DBG-DUMPHEADR Parameter",
                   MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap Accounting")
                Exit Sub
            End Try
            log.Debug("dumpHeader - " & "dumping to " & dumpHd)
            ' clear the Header
            If CStr(aDWS.Cells(1, 1).Value) <> "" Then
                aRange = aDWS.Range(aDWS.Cells(1, 1), aDWS.Cells(1000, 1))
                aRange.EntireRow.Delete()
            End If
            ' dump the Header
            Dim aTStrRec As New SAPCommon.TStrRec
            Dim aFieldArray() As String = {}
            Dim aValueArray() As String = {}
            For Each aTStrRec In aHdrRec.aTDataRecCol
                Array.Resize(aFieldArray, aFieldArray.Length + 1)
                aFieldArray(aFieldArray.Length - 1) = aTStrRec.getKey()
                Array.Resize(aValueArray, aValueArray.Length + 1)
                aValueArray(aValueArray.Length - 1) = aTStrRec.formated()
            Next
            aRange = aDWS.Range(aDWS.Cells(1, 1), aDWS.Cells(1, aFieldArray.Length))
            aRange.Value = aFieldArray
            aRange = aDWS.Range(aDWS.Cells(2, 1), aDWS.Cells(2, aValueArray.Length))
            aRange.Value = aValueArray
        End If
    End Sub

    Public Sub dumpData()
        Dim dumpDt As String = If(aIntPar.value("DBG", "DUMPDATA") <> "", aIntPar.value("DBG", "DUMPDATA"), "")
        If dumpDt <> "" Then
            Dim aDWS As Excel.Worksheet
            Dim aWB As Excel.Workbook
            Dim aRange As Excel.Range
            aWB = Globals.SapMmMdExcelAddin.Application.ActiveWorkbook
            Try
                aDWS = aWB.Worksheets(dumpDt)
                aDWS.Activate()
            Catch Exc As System.Exception
                log.Warn("dumpData - " & "No " & dumpDt & " Sheet in current workbook.")
                MsgBox("No " & dumpDt & " Sheet in current workbook. Check the DBG-DUMPDATA Parameter",
                       MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap Accounting")
                Exit Sub
            End Try
            log.Debug("dumpData - " & "dumping to " & dumpDt)
            ' clear the Data
            If CStr(aDWS.Cells(5, 1).Value) <> "" Then
                aRange = aDWS.Range(aDWS.Cells(5, 1), aDWS.Cells(1000, 1))
                aRange.EntireRow.Delete()
            End If

            Dim aKvB_Dic As KeyValuePair(Of String, TData)
            Dim aKvB_Rec As KeyValuePair(Of String, TDataRec)
            Dim aData As TData
            Dim aData_Am As New TData(aIntPar)
            Dim aDataRec As New TDataRec(aIntPar)
            Dim aDataRec_Am As New TDataRec(aIntPar)
            Dim i As Int64
            Dim aTStrRec As New SAPCommon.TStrRec
            i = 6
            For Each aKvB_Dic In aDataDic.aTDataDic
                aData = aKvB_Dic.Value
                aDWS.Cells(i, 1).Value = aKvB_Dic.Key
                For Each aKvB_Rec In aData.aTDataDic
                    aDataRec = aKvB_Rec.Value
                    Dim aFieldArray() As String = {}
                    Dim aValueArray() As String = {}
                    For Each aTStrRec In aDataRec.aTDataRecCol
                        Array.Resize(aFieldArray, aFieldArray.Length + 1)
                        aFieldArray(aFieldArray.Length - 1) = aTStrRec.getKey()
                        Array.Resize(aValueArray, aValueArray.Length + 1)
                        aValueArray(aValueArray.Length - 1) = aTStrRec.formated()
                    Next
                    aRange = aDWS.Range(aDWS.Cells(i, 1), aDWS.Cells(i, aFieldArray.Length))
                    aRange.Value = aFieldArray
                    aRange = aDWS.Range(aDWS.Cells(i + 1, 1), aDWS.Cells(i + 1, aValueArray.Length))
                    aRange.Value = aValueArray
                    i += 2
                Next
                i += 2
            Next
        End If
    End Sub

    Public Sub ws_output_line(pStructure As String, pFieldArray() As String, pIsValueArray() As String, ByRef pWs As Excel.Worksheet, pDataKey As String, i As UInt64, jMax As UInt64, Optional pCoff As Integer = 0, Optional ByRef pClear As Boolean = False, Optional pKey As String = "")
        Dim aTData As New TData(pPar:=aIntPar)
        If aDataDic.aTDataDic.ContainsKey(pStructure) Then
            aTData = aDataDic.aTDataDic(pStructure)
            aTData.setFieldArray(pFieldArray:=pFieldArray, pIsValueArray:=pIsValueArray)
            aTData.ws_output_line(pWs:=pWs, pDataKey:="", i:=i, jMax:=jMax, pClear:=pClear, pKey:=pKey)
        End If
    End Sub

    Public Function ws_output(pStructure As String, pFieldArray() As String, pIsValueArray() As String, ByRef pWs As Excel.Worksheet, pDataKey As String, i As UInt64, jMax As UInt64, Optional pCoff As Integer = 0, Optional ByRef pClear As Boolean = False, Optional pKey As String = "") As UInt64
        Dim aTData As New TData(pPar:=aIntPar)
        If aDataDic.aTDataDic.ContainsKey(pStructure) Then
            aTData = aDataDic.aTDataDic(pStructure)
            aTData.setFieldArray(pFieldArray:=pFieldArray, pIsValueArray:=pIsValueArray)
            ws_output = aTData.ws_output(pWs:=pWs, pDataKey:="", i:=i, jMax:=jMax, pClear:=pClear, pKey:=pKey)
        Else
            ws_output = 0
        End If
    End Function

    Public Function getField(pStrucName As String, pFieldname As String) As SAPCommon.TStrRec
        Dim aTStrRec As SAPCommon.TStrRec
        Try
            getField = aDataDic.aTDataDic(pStrucName).getFirstRecord().getColumn(pStrucName & "-" & pFieldname)
        Catch ex As Exception
            aTStrRec = Nothing
        End Try
    End Function

    Public Function getField(pName As String) As SAPCommon.TStrRec
        Dim aNameArray() As String
        Dim aSTRUCNAME As String = ""
        Dim aFIELDNAME As String = ""
        If InStr(pName, "-") <> 0 Then
            aNameArray = Split(pName, "-")
            aSTRUCNAME = aNameArray(0)
            aFIELDNAME = aNameArray(1)
        Else
            aSTRUCNAME = ""
            aFIELDNAME = pName
        End If
        Try
            If String.IsNullOrEmpty(aSTRUCNAME) Then
                getField = aHdrRec.aTDataRecCol(aFIELDNAME).getColumn(aSTRUCNAME & "-" & aFIELDNAME)
            Else
                getField = aDataDic.aTDataDic(aSTRUCNAME).getFirstRecord().getColumn(aSTRUCNAME & "-" & aFIELDNAME)
            End If
        Catch ex As Exception
            getField = Nothing
        End Try
    End Function

    Public Function getMaintStat() As String
        Dim TStrRec As SAPCommon.TStrRec = getField(pStrucName:="CLIENTDATA", pFieldname:="MAINT_STAT")
        getMaintStat = If(TStrRec Is Nothing, "", TStrRec.Value)
    End Function
End Class
