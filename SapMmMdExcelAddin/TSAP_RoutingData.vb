' Copyright 2023 Hermann Mundprecht
' This file is licensed under the terms of the license 'CC BY 4.0'. 
' For a human readable version of the license, see https://creativecommons.org/licenses/by/4.0/

Imports SAP.Middleware.Connector

Public Class TSAP_RoutingData

    Public aHdrRec As TDataRec
    Public aDataDic As TDataDic

    Public aStrucDic As Dictionary(Of String, RfcStructureMetadata)
    Public aParamDic As Dictionary(Of String, RfcParameterMetadata)
    '   Public aFieldsDic As Dictionary(Of String, String())

    Private aPar As SAPCommon.TStr
    Private aIntPar As SAPCommon.TStr
    Private aFunction As String
    Private Shared ReadOnly log As log4net.ILog = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType)

    Public Sub New(ByRef pPar As SAPCommon.TStr, ByRef pIntPar As SAPCommon.TStr, ByRef pSAPRouting As SAPRouting, pFunction As String)
        aPar = pPar
        aIntPar = pIntPar
        aFunction = pFunction
        aDataDic = New TDataDic(aIntPar)
        aHdrRec = New TDataRec(aIntPar)
        aStrucDic = New Dictionary(Of String, RfcStructureMetadata)
        aParamDic = New Dictionary(Of String, RfcParameterMetadata)
        '       aFieldsDic = New Dictionary(Of String, String())
        ' get Metadata
        If pFunction = "create" Then
            pSAPRouting.getMeta_Create(aParamDic, aStrucDic)
        ElseIf pFunction = "change" Then
            pSAPRouting.getMeta_Change(aParamDic, aStrucDic)
        End If
    End Sub

    Public Function fillHeader(pData As TData) As Boolean
        Dim aKvb As KeyValuePair(Of String, SAPCommon.TStrRec)
        Dim aTStrRec As SAPCommon.TStrRec
        Dim aNewTStrRec As SAPCommon.TStrRec
        Dim aNewHdrRec As New TDataRec(aIntPar)
        Dim aStrucName() As String
        Dim aLen As Integer = 0
        For Each aKvb In aPar.getData()
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
                If valid_StructureX_Field(aNewTStrRec, aLen) Then
                    If aLen = 1 Then
                        aNewHdrRec.setValues(aNewTStrRec.Strucname & "_X-" & aNewTStrRec.Fieldname, "X", "", "")
                    Else
                        aNewHdrRec.setValues(aNewTStrRec.Strucname & "_X-" & aNewTStrRec.Fieldname, aNewTStrRec.Value, aNewTStrRec.Currency, aNewTStrRec.Format, pEmptyChar:="")
                    End If
                End If
            Next
        Next
        ' First fill the value from the paramters and tehn overwrite them from the posting record
        Dim aPostRec As New TDataRec(aIntPar)
        aPostRec = pData.getFirstRecord()
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
                If valid_StructureX_Field(aNewTStrRec, aLen) Then
                    If aLen = 1 Then
                        aNewHdrRec.setValues(aNewTStrRec.Strucname & "_X-" & aNewTStrRec.Fieldname, "X", "", "")
                    Else
                        aNewHdrRec.setValues(aNewTStrRec.Strucname & "_X-" & aNewTStrRec.Fieldname, aNewTStrRec.Value, aNewTStrRec.Currency, aNewTStrRec.Format, pEmptyChar:="")
                    End If
                End If
            Next
        Next
        aHdrRec = aNewHdrRec
        fillHeader = True
    End Function

    Public Function fillData(pData As TData) As Boolean
        Dim aKvB As KeyValuePair(Of String, TDataRec)
        Dim aTDataRec As TDataRec
        Dim aTStrRec As SAPCommon.TStrRec
        Dim aCnt As UInt64
        aDataDic = New TDataDic(aIntPar)
        fillData = True
        aCnt = 1
        For Each aKvB In pData.aTDataDic
            aTDataRec = aKvB.Value
            For Each aTStrRec In aTDataRec.aTDataRecCol
                If valid_Table_Field(aTStrRec) Then
                    Dim aStrucName() As String = Split(aTStrRec.Strucname, "+")
                    For s As Integer = 0 To aStrucName.Length - 1
                        Dim aNewTStrRec As SAPCommon.TStrRec = New SAPCommon.TStrRec(aStrucName(s), aTStrRec.Fieldname, aTStrRec.Value, aTStrRec.Currency, aTStrRec.Format)
                        aDataDic.addValue(CStr(aCnt), aNewTStrRec)
                    Next
                End If
                Dim aLen As Integer
                If valid_TableX_Field(aTStrRec, aLen) Then
                    Dim aStrucName() As String = Split(aTStrRec.Strucname, "+")
                    For s As Integer = 0 To aStrucName.Length - 1
                        Dim aNewTStrRec As SAPCommon.TStrRec
                        If aLen = 1 Then
                            aNewTStrRec = New SAPCommon.TStrRec(aStrucName(s) & "_X", aTStrRec.Fieldname, "X", "", "")
                        Else
                            aNewTStrRec = New SAPCommon.TStrRec(aStrucName(s) & "_X", aTStrRec.Fieldname, aTStrRec.Value, aTStrRec.Currency, aTStrRec.Format)
                        End If
                        aDataDic.addValue(CStr(aCnt), aNewTStrRec)
                    Next
                End If
            Next
            aCnt += 1
        Next
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

    Public Function valid_StructureX_Field(pTStrRec As SAPCommon.TStrRec, ByRef pLen As Integer) As Boolean
        Dim aStrucName() As String
        valid_StructureX_Field = False
        aStrucName = Split(pTStrRec.Strucname, "+")
        For s As Integer = 0 To aStrucName.Length - 1
            If aStrucDic.ContainsKey("S|" & aStrucName(s) & "_X") Then
                valid_StructureX_Field = isInStructure(pTStrRec.Fieldname, aStrucDic("S|" & aStrucName(s) & "_X"), pLen)
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

    Public Function valid_TableX_Field(pTStrRec As SAPCommon.TStrRec, ByRef pLen As Integer) As Boolean
        Dim aStrucName() As String
        valid_TableX_Field = False
        aStrucName = Split(pTStrRec.Strucname, "+")
        For s As Integer = 0 To aStrucName.Length - 1
            If aStrucDic.ContainsKey("T|" & aStrucName(s) & "_X") Then
                valid_TableX_Field = isInStructure(pTStrRec.Fieldname, aStrucDic("T|" & aStrucName(s) & "_X"), pLen)
            End If
        Next
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

    Public Function getHdrField(pStrucName As String, pFieldname As String) As SAPCommon.TStrRec
        Try
            getHdrField = aHdrRec.aTDataRecCol(pStrucName & "-" & pFieldname)
        Catch ex As Exception
            getHdrField = Nothing
        End Try
    End Function

    Public Function getDataField(pStrucName As String, pFieldname As String) As SAPCommon.TStrRec
        Try
            getDataField = aDataDic.aTDataDic(pStrucName).getFirstRecord().getColumn(pStrucName & "-" & pFieldname)
        Catch ex As Exception
            getDataField = Nothing
        End Try
    End Function

    Public Function getHdrField(pName As String) As SAPCommon.TStrRec
        Try
            getHdrField = aHdrRec.aTDataRecCol(pName)
        Catch ex As Exception
            getHdrField = Nothing
        End Try
    End Function

    Public Function getDataField(pName As String) As SAPCommon.TStrRec
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
            getDataField = aDataDic.aTDataDic(aSTRUCNAME).getFirstRecord().getColumn(aSTRUCNAME & "-" & aFIELDNAME)
        Catch ex As Exception
            getDataField = Nothing
        End Try
    End Function

    Public Function getMaintStat() As String
        Dim TStrRec As SAPCommon.TStrRec = getHdrField(pStrucName:="CLIENTDATA", pFieldname:="MAINT_STAT")
        getMaintStat = If(TStrRec Is Nothing, "", TStrRec.Value)
    End Function

    Public Function getMaterial() As String
        ' try the HEADDATA from BAPI_MATERIAL_SAVEDATA
        Dim aTStrRec As SAPCommon.TStrRec = getMaterialRec()
        getMaterial = If(aTStrRec Is Nothing, "", aTStrRec.Value)
    End Function

    Public Function getMaterialRec() As SAPCommon.TStrRec
        ' try the HEADDATA from BAPI_MATERIAL_SAVEDATA
        getMaterialRec = getHdrField(pStrucName:="HEADDATA", pFieldname:="MATERIAL")
        If getMaterialRec Is Nothing Then
            getMaterialRec = getHdrField(pStrucName:="HEADDATA", pFieldname:="MATERIAL_EXTERNAL")
        End If
        If getMaterialRec Is Nothing Then
            getMaterialRec = getHdrField(pStrucName:="HEADDATA", pFieldname:="MATERIAL_LONG")
        End If
        ' try the Import Fields from BAPI_MATERIAL_GET_ALL
        If getMaterialRec Is Nothing Then
            getMaterialRec = getHdrField(pStrucName:="", pFieldname:="MATERIAL")
        End If
        If getMaterialRec Is Nothing Then
            getMaterialRec = getHdrField(pStrucName:="", pFieldname:="MATERIAL_LONG")
        End If
    End Function

End Class
