' Copyright 2022 Hermann Mundprecht
' This file is licensed under the terms of the license 'CC BY 4.0'. 
' For a human readable version of the license, see https://creativecommons.org/licenses/by/4.0/

Imports SAP.Middleware.Connector

Public Class SAPMaterial

    Private Shared ReadOnly log As log4net.ILog = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType)
    Private oRfcFunction As IRfcFunction
    Private destination As RfcCustomDestination
    Private sapcon As SapCon

    Sub New(aSapCon As SapCon)
        Try
            log.Debug("New - " & "checking connection")
            sapcon = aSapCon
            aSapCon.getDestination(destination)
            sapcon.checkCon()
        Catch ex As System.Exception
            log.Error("New - Exception=" & ex.ToString)
            MsgBox("New failed! " & ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAPMaterial")
        End Try
    End Sub

    Private Sub addToStrucDic(pArrayName As String, pRfcStructureMetadata As RfcStructureMetadata, ByRef pStrucDic As Dictionary(Of String, RfcStructureMetadata))
        If pStrucDic.ContainsKey(pArrayName) Then
            pStrucDic.Remove(pArrayName)
            pStrucDic.Add(pArrayName, pRfcStructureMetadata)
        Else
            pStrucDic.Add(pArrayName, pRfcStructureMetadata)
        End If
    End Sub

    Private Sub addToFieldDic(pArrayName As String, pRfcStructureMetadata As RfcParameterMetadata, ByRef pFieldDic As Dictionary(Of String, RfcParameterMetadata))
        If pFieldDic.ContainsKey(pArrayName) Then
            pFieldDic.Remove(pArrayName)
            pFieldDic.Add(pArrayName, pRfcStructureMetadata)
        Else
            pFieldDic.Add(pArrayName, pRfcStructureMetadata)
        End If
    End Sub

    Public Sub getMeta_Change(ByRef pFieldDic As Dictionary(Of String, RfcParameterMetadata), ByRef pStrucDic As Dictionary(Of String, RfcStructureMetadata))
        Dim aStructures As String() = {"HEADDATA", "CLIENTDATA", "CLIENTDATAX", "PLANTDATA", "PLANTDATAX", "FORECASTPARAMETERS", "FORECASTPARAMETERSX", "PLANNINGDATA", "PLANNINGDATAX", "STORAGELOCATIONDATA", "STORAGELOCATIONDATAX", "VALUATIONDATA", "VALUATIONDATAX", "WAREHOUSENUMBERDATA", "WAREHOUSENUMBERDATAX", "SALESDATA", "SALESDATAX", "STORAGETYPEDATA", "STORAGETYPEDATAX"}
        Dim aImports As String() = {"FLAG_ONLINE", "FLAG_CAD_CALL", "NO_DEQUEUE", "NO_ROLLBACK_WORK"}
        Dim aTables As String() = {"MATERIALDESCRIPTION", "UNITSOFMEASURE", "UNITSOFMEASUREX", "INTERNATIONALARTNOS", "MATERIALLONGTEXT", "TAXCLASSIFICATIONS", "RETURNMESSAGES", "PRTDATA", "PRTDATAX", "EXTENSIONIN", "EXTENSIONINX"}
        Try
            log.Debug("getMeta_Change - " & "creating Function BAPI_MATERIAL_SAVEDATA")
            oRfcFunction = destination.Repository.CreateFunction("BAPI_MATERIAL_SAVEDATA")
            Dim oStructure As IRfcStructure
            Dim oTable As IRfcTable
            ' Imports
            For s As Integer = 0 To aImports.Length - 1
                addToFieldDic("I|" & aImports(s), oRfcFunction.Metadata.Item(aImports(s)), pFieldDic)
            Next
            ' Import Strcutures
            For s As Integer = 0 To aStructures.Length - 1
                oStructure = oRfcFunction.GetStructure(aStructures(s))
                addToStrucDic("S|" & aStructures(s), oStructure.Metadata, pStrucDic)
            Next
            For s As Integer = 0 To aTables.Length - 1
                oTable = oRfcFunction.GetTable(aTables(s))
                addToStrucDic("T|" & aTables(s), oTable.Metadata.LineType, pStrucDic)
            Next
        Catch Ex As System.Exception
            log.Error("getMeta_Change - Exception=" & Ex.ToString)
            MsgBox("Error: Exception " & Ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAPMaterial")
        Finally
            log.Debug("getMeta_GetDetail - " & "EndContext")
            RfcSessionManager.EndContext(destination)
        End Try
    End Sub

    Public Sub getMeta_GetAll(ByRef pFieldDic As Dictionary(Of String, RfcParameterMetadata), ByRef pStrucDic As Dictionary(Of String, RfcStructureMetadata))
        Dim aStructures As String() = {"CLIENTDATA", "PLANTDATA", "FORECASTPARAMETERS", "PLANNINGDATA", "STORAGELOCATIONDATA", "VALUATIONDATA", "WAREHOUSENUMBERDATA", "SALESDATA", "STORAGETYPEDATA", "PRTDATA", "LIFOVALUATIONDATA"}
        Dim aImports As String() = {"MATERIAL", "COMP_CODE", "VAL_AREA", "VAL_TYPE", "PLANT", "STGE_LOC", "SALESORG", "DISTR_CHAN", "WHSENUMBER", "STGE_TYPE", "LIFO_VALUATION_LEVEL", "KZRFB_ALL", "MATERIAL_LONG"}
        Dim aTables As String() = {"MATERIALDESCRIPTION", "UNITSOFMEASURE", "INTERNATIONALARTNOS", "MATERIALLONGTEXT", "TAXCLASSIFICATIONS", "EXTENSIONOUT", "RETURN"}
        Try
            log.Debug("getMeta_GetAll - " & "creating Function BAPI_MATERIAL_GET_ALL")
            oRfcFunction = destination.Repository.CreateFunction("BAPI_MATERIAL_GET_ALL")
            Dim oStructure As IRfcStructure
            Dim oTable As IRfcTable
            ' Imports
            For s As Integer = 0 To aImports.Length - 1
                addToFieldDic("I|" & aImports(s), oRfcFunction.Metadata.Item(aImports(s)), pFieldDic)
            Next
            ' Import Strcutures
            For s As Integer = 0 To aStructures.Length - 1
                oStructure = oRfcFunction.GetStructure(aStructures(s))
                addToStrucDic("S|" & aStructures(s), oStructure.Metadata, pStrucDic)
            Next
            For s As Integer = 0 To aTables.Length - 1
                oTable = oRfcFunction.GetTable(aTables(s))
                addToStrucDic("T|" & aTables(s), oTable.Metadata.LineType, pStrucDic)
            Next
        Catch Ex As System.Exception
            log.Error("getMeta_GetAll - Exception=" & Ex.ToString)
            MsgBox("Error: Exception " & Ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAPMaterial")
        Finally
            log.Debug("getMeta_GetAll - " & "EndContext")
            RfcSessionManager.EndContext(destination)
        End Try
    End Sub

    Public Function Change(pData As TSAP_MatData, Optional pOKMsg As String = "OK", Optional pCheck As Boolean = False) As String
        Change = ""
        Try
            oRfcFunction = destination.Repository.CreateFunction("BAPI_MATERIAL_SAVEDATA")
            RfcSessionManager.BeginContext(destination)
            Dim oRETURN As IRfcTable = oRfcFunction.GetTable("RETURNMESSAGES")
            Dim oMATERIALDESCRIPTION As IRfcTable = oRfcFunction.GetTable("MATERIALDESCRIPTION")
            Dim oUNITSOFMEASURE As IRfcTable = oRfcFunction.GetTable("UNITSOFMEASURE")
            Dim oUNITSOFMEASUREX As IRfcTable = oRfcFunction.GetTable("UNITSOFMEASUREX")
            Dim oINTERNATIONALARTNOS As IRfcTable = oRfcFunction.GetTable("INTERNATIONALARTNOS")
            Dim oMATERIALLONGTEXT As IRfcTable = oRfcFunction.GetTable("MATERIALLONGTEXT")
            Dim oTAXCLASSIFICATIONS As IRfcTable = oRfcFunction.GetTable("TAXCLASSIFICATIONS")
            Dim oPRTDATA As IRfcTable = oRfcFunction.GetTable("PRTDATA")
            Dim oPRTDATAX As IRfcTable = oRfcFunction.GetTable("PRTDATAX")
            Dim oEXTENSIONIN As IRfcTable = oRfcFunction.GetTable("EXTENSIONIN")
            Dim oEXTENSIONINX As IRfcTable = oRfcFunction.GetTable("EXTENSIONINX")
            oMATERIALDESCRIPTION.Clear()
            oUNITSOFMEASURE.Clear()
            oUNITSOFMEASUREX.Clear()
            oINTERNATIONALARTNOS.Clear()
            oMATERIALLONGTEXT.Clear()
            oTAXCLASSIFICATIONS.Clear()
            oPRTDATA.Clear()
            oPRTDATAX.Clear()
            oEXTENSIONIN.Clear()
            oEXTENSIONINX.Clear()
            oRETURN.Clear()

            Dim aTStrRec As SAPCommon.TStrRec
            Dim oStruc As IRfcStructure
            ' set the header values
            For Each aTStrRec In pData.aHdrRec.aTDataRecCol
                If aTStrRec.Strucname <> "" Then
                    oStruc = oRfcFunction.GetStructure(aTStrRec.Strucname)
                    oStruc.SetValue(aTStrRec.Fieldname, aTStrRec.formated)
                Else
                    oRfcFunction.SetValue(aTStrRec.Fieldname, aTStrRec.formated)
                End If
            Next
            ' set the table fields
            pData.aDataDic.to_IRfcTable(pKey:="MATERIALDESCRIPTION", pIRfcTable:=oMATERIALDESCRIPTION)
            pData.aDataDic.to_IRfcTable(pKey:="UNITSOFMEASURE", pIRfcTable:=oUNITSOFMEASURE)
            pData.aDataDic.to_IRfcTableX(pKey:="UNITSOFMEASURE", pIRfcTable:=oUNITSOFMEASUREX, pKeyFields:={"ALT_UNIT", "ALT_UNIT_ISO"})
            pData.aDataDic.to_IRfcTable(pKey:="INTERNATIONALARTNOS", pIRfcTable:=oINTERNATIONALARTNOS)
            pData.aDataDic.to_IRfcTable(pKey:="MATERIALLONGTEXT", pIRfcTable:=oMATERIALLONGTEXT)
            pData.aDataDic.to_IRfcTable(pKey:="TAXCLASSIFICATIONS", pIRfcTable:=oTAXCLASSIFICATIONS)
            pData.aDataDic.to_IRfcTable(pKey:="PRTDATA", pIRfcTable:=oPRTDATA)
            pData.aDataDic.to_IRfcTableX(pKey:="PRTDATA", pIRfcTable:=oPRTDATAX, pKeyFields:={"PLANT"})
            pData.aDataDic.to_IRfcTable(pKey:="EXTENSIONIN", pIRfcTable:=oEXTENSIONIN)
            pData.aDataDic.to_IRfcTable(pKey:="EXTENSIONINX", pIRfcTable:=oEXTENSIONINX)
            ' call the BAPI
            oRfcFunction.Invoke(destination)
            Dim sRETURN As IRfcStructure = oRfcFunction.GetStructure("RETURN")
            Dim aErr As Boolean = False
            If sRETURN.GetValue("TYPE") = "S" Then
                Change = Change & ";" & sRETURN.GetValue("MESSAGE")
            Else
                For i As Integer = 0 To oRETURN.Count - 1
                    If oRETURN(i).GetValue("TYPE") <> "I" And oRETURN(i).GetValue("TYPE") <> "W" Then
                        Change = Change & ";" & oRETURN(i).GetValue("MESSAGE")
                        If oRETURN(i).GetValue("TYPE") <> "S" And oRETURN(i).GetValue("TYPE") <> "W" Then
                            aErr = True
                        End If
                    End If
                Next i
            End If
            If aErr = False Then
                Dim aSAPBapiTranctionCommit As New SAPBapiTranctionCommit(sapcon)
                aSAPBapiTranctionCommit.commit(pWait:="X")
            End If
            Change = If(Change = "", pOKMsg, If(aErr = False, pOKMsg & Change, "Error" & Change))

        Catch Ex As System.Exception
            MsgBox("Error: Exception " & Ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAPMaterial")
            Change = "Error: Exception in Change"
        Finally
            RfcSessionManager.EndContext(destination)
        End Try
    End Function

    Public Function GetAll(pData As TSAP_MatData, Optional pOKMsg As String = "OK", Optional pCheck As Boolean = False) As String
        GetAll = ""
        Try
            oRfcFunction = destination.Repository.CreateFunction("BAPI_MATERIAL_GET_ALL")
            RfcSessionManager.BeginContext(destination)
            Dim oRETURN As IRfcTable = oRfcFunction.GetTable("RETURN")
            Dim oMATERIALDESCRIPTION As IRfcTable = oRfcFunction.GetTable("MATERIALDESCRIPTION")
            Dim oUNITSOFMEASURE As IRfcTable = oRfcFunction.GetTable("UNITSOFMEASURE")
            Dim oINTERNATIONALARTNOS As IRfcTable = oRfcFunction.GetTable("INTERNATIONALARTNOS")
            Dim oMATERIALLONGTEXT As IRfcTable = oRfcFunction.GetTable("MATERIALLONGTEXT")
            Dim oTAXCLASSIFICATIONS As IRfcTable = oRfcFunction.GetTable("TAXCLASSIFICATIONS")
            Dim oEXTENSIONOUT As IRfcTable = oRfcFunction.GetTable("EXTENSIONOUT")
            '            Dim oNFMCHARGEWEIGHTS As IRfcTable = oRfcFunction.GetTable("NFMCHARGEWEIGHTS")
            '           Dim oNFMSTRUCTURALWEIGHTS As IRfcTable = oRfcFunction.GetTable("NFMSTRUCTURALWEIGHTS")
            oMATERIALDESCRIPTION.Clear()
            oUNITSOFMEASURE.Clear()
            oINTERNATIONALARTNOS.Clear()
            oMATERIALLONGTEXT.Clear()
            oTAXCLASSIFICATIONS.Clear()
            '            oNFMCHARGEWEIGHTS.Clear()
            '            oNFMSTRUCTURALWEIGHTS.Clear()
            oEXTENSIONOUT.Clear()
            oRETURN.Clear()

            Dim aTStrRec As SAPCommon.TStrRec
            Dim oStruc As IRfcStructure
            ' set the header values
            For Each aTStrRec In pData.aHdrRec.aTDataRecCol
                If aTStrRec.Strucname <> "" Then
                    oStruc = oRfcFunction.GetStructure(aTStrRec.Strucname)
                    oStruc.SetValue(aTStrRec.Fieldname, aTStrRec.formated)
                Else
                    oRfcFunction.SetValue(aTStrRec.Fieldname, aTStrRec.formated)
                End If
            Next
            ' call the BAPI
            oRfcFunction.Invoke(destination)

            Dim aErr As Boolean = False
            For i As Integer = 0 To oRETURN.Count - 1
                If oRETURN(i).GetValue("TYPE") <> "I" And oRETURN(i).GetValue("TYPE") <> "W" Then
                    GetAll = GetAll & ";" & oRETURN(i).GetValue("MESSAGE")
                    If oRETURN(i).GetValue("TYPE") <> "S" And oRETURN(i).GetValue("TYPE") <> "W" Then
                        aErr = True
                    End If
                End If
            Next i
            GetAll = If(GetAll = "", pOKMsg, If(aErr = False, pOKMsg & GetAll, "Error" & GetAll))

            If aErr = False Then
                ' process the return structures
                pData.aDataDic.addValue("0", oStruc:=oRfcFunction.GetStructure("CLIENTDATA"), pStrucName:="CLIENTDATA")
                pData.aDataDic.addValue("0", oStruc:=oRfcFunction.GetStructure("PLANTDATA"), pStrucName:="PLANTDATA")
                pData.aDataDic.addValue("0", oStruc:=oRfcFunction.GetStructure("FORECASTPARAMETERS"), pStrucName:="FORECASTPARAMETERS")
                pData.aDataDic.addValue("0", oStruc:=oRfcFunction.GetStructure("PLANNINGDATA"), pStrucName:="PLANNINGDATA")
                pData.aDataDic.addValue("0", oStruc:=oRfcFunction.GetStructure("STORAGELOCATIONDATA"), pStrucName:="STORAGELOCATIONDATA")
                pData.aDataDic.addValue("0", oStruc:=oRfcFunction.GetStructure("VALUATIONDATA"), pStrucName:="VALUATIONDATA")
                pData.aDataDic.addValue("0", oStruc:=oRfcFunction.GetStructure("WAREHOUSENUMBERDATA"), pStrucName:="WAREHOUSENUMBERDATA")
                pData.aDataDic.addValue("0", oStruc:=oRfcFunction.GetStructure("SALESDATA"), pStrucName:="SALESDATA")
                pData.aDataDic.addValue("0", oStruc:=oRfcFunction.GetStructure("STORAGETYPEDATA"), pStrucName:="STORAGETYPEDATA")
                pData.aDataDic.addValue("0", oStruc:=oRfcFunction.GetStructure("PRTDATA"), pStrucName:="PRTDATA")
                pData.aDataDic.addValue("0", oStruc:=oRfcFunction.GetStructure("LIFOVALUATIONDATA"), pStrucName:="LIFOVALUATIONDATA")
                ' process the return tables
                pData.aDataDic.addValues(oTable:=oMATERIALDESCRIPTION, pStrucName:="MATERIALDESCRIPTION")
                pData.aDataDic.addValues(oTable:=oUNITSOFMEASURE, pStrucName:="UNITSOFMEASURE")
                pData.aDataDic.addValues(oTable:=oINTERNATIONALARTNOS, pStrucName:="INTERNATIONALARTNOS")
                pData.aDataDic.addValues(oTable:=oMATERIALLONGTEXT, pStrucName:="MATERIALLONGTEXT")
                pData.aDataDic.addValues(oTable:=oTAXCLASSIFICATIONS, pStrucName:="TAXCLASSIFICATIONS")
                '                pData.aDataDic.addValues(oTable:=oNFMCHARGEWEIGHTS, pStrucName:="NFMCHARGEWEIGHTS")
                '                pData.aDataDic.addValues(oTable:=oNFMSTRUCTURALWEIGHTS, pStrucName:="NFMSTRUCTURALWEIGHTS")
                pData.aDataDic.addValues(oTable:=oEXTENSIONOUT, pStrucName:="EXTENSIONOUT")
            End If
        Catch Ex As System.Exception
            MsgBox("Error: Exception " & Ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAPMaterial")
            GetAll = "Error: Exception in GetAll"
        Finally
            RfcSessionManager.EndContext(destination)
        End Try
    End Function

End Class
