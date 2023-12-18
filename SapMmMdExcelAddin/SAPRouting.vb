' Copyright 2022 Hermann Mundprecht
' This file is licensed under the terms of the license 'CC BY 4.0'. 
' For a human readable version of the license, see https://creativecommons.org/licenses/by/4.0/

Imports SAP.Middleware.Connector

Public Class SAPRouting

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
            MsgBox("New failed! " & ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAPRouting")
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

    Public Sub getMeta_Create(ByRef pFieldDic As Dictionary(Of String, RfcParameterMetadata), ByRef pStrucDic As Dictionary(Of String, RfcStructureMetadata))
        Dim aStructures As String() = {}
        Dim aImports As String() = {"TESTRUN", "PROFILE", "BOMUSAGE", "APPLICATION"}
        Dim aTables As String() = {"TASK", "MATERIALTASKALLOCATION", "SEQUENCE", "OPERATION", "SUBOPERATION", "REFERENCEOPERATION", "WORKCENTERREFERENCE", "COMPONENTALLOCATION", "PRODUCTIONRESOURCE", "INSPCHARACTERISTIC", "TEXTALLOCATION", "TEXT", "RETURN", "TASK_SEGMENT", "DEPENDENCY_ALLOCATION", "DEPENDENCY_ORDER", "DEPENDENCY_DATA", "DEPENDENCY_DESCRIPTION", "DEPENDENCY_DOCUMENTATION", "DEPENDENCY_SOURCE"}
        Try
            log.Debug("getMeta_Change - " & "creating Function BAPI_ROUTING_CREATE")
            oRfcFunction = destination.Repository.CreateFunction("BAPI_ROUTING_CREATE")
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
            MsgBox("Error: Exception " & Ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAPRouting")
        Finally
            log.Debug("getMeta_GetDetail - " & "EndContext")
            RfcSessionManager.EndContext(destination)
        End Try
    End Sub

    Public Sub getMeta_Change(ByRef pFieldDic As Dictionary(Of String, RfcParameterMetadata), ByRef pStrucDic As Dictionary(Of String, RfcStructureMetadata))
        Dim aStructures As String() = {"TASK", "TASK_X"}
        Dim aImports As String() = {"CHANGE_NO", "KEY_DATE", "TASK_LIST_GROUP", "GROUP_COUNTER", "MATERIAL", "PLANT"}
        Dim aTables As String() = {"MATERIAL_TASK_ALLOCATIONS", "MATERIAL_TASK_ALLOCATIONS_X", "SEQUENCES", "SEQUENCES_X", "OPERATIONS", "OPERATIONS_X", "SUBOPERATIONS", "SUBOPERATIONS_X", "PRODUCTION_RESOURCES_TOOLS", "PRODUCTION_RESOURCES_TOOLS_X", "COMPONENT_ALLOCATIONS", "COMPONENT_ALLOCATIONS_X", "INSPECTION_CHARACTERISTICS", "INSPECTION_CHARACTERISTICS_X", "INSPECTION_VALUES", "INSPECTION_VALUES_X", "REFERENCED_OPERATIONS", "TEXT", "TEXT_ALLOCATIONS", "RETURN", "SEGMENT_TASK_MAINTAIN"}
        Try
            log.Debug("getMeta_Change - " & "creating Function ROUTING_MAINTAIN")
            oRfcFunction = destination.Repository.CreateFunction("ROUTING_MAINTAIN")
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
            MsgBox("Error: Exception " & Ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAPRouting")
        Finally
            log.Debug("getMeta_GetDetail - " & "EndContext")
            RfcSessionManager.EndContext(destination)
        End Try
    End Sub

    Public Function Create(pData As TSAP_RoutingData, Optional pOKMsg As String = "OK", Optional pCheck As Boolean = False) As String
        Create = ""
        Try
            oRfcFunction = destination.Repository.CreateFunction("BAPI_ROUTING_CREATE")
            RfcSessionManager.BeginContext(destination)
            Dim oRETURN As IRfcTable = oRfcFunction.GetTable("RETURN")
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
            For Each aKey In pData.aDataDic.aTDataDic.Keys
                Dim oTable As IRfcTable = oRfcFunction.GetTable(aKey)
                oTable.Clear()
                pData.aDataDic.to_IRfcTable(pKey:=aKey, pIRfcTable:=oTable)
            Next
            ' call the BAPI
            oRfcFunction.Invoke(destination)
            Dim aErr As Boolean = False
            For i As Integer = 0 To oRETURN.Count - 1
                Create = Create & ";" & oRETURN(i).GetValue("MESSAGE")
                If oRETURN(i).GetValue("TYPE") <> "S" And oRETURN(i).GetValue("TYPE") <> "I" And oRETURN(i).GetValue("TYPE") <> "W" Then
                    aErr = True
                End If
            Next
            If aErr = False Then
                Dim aSAPBapiTranctionCommit As New SAPBapiTranctionCommit(sapcon)
                aSAPBapiTranctionCommit.commit(pWait:="X")
            End If
            Create = If(Create = "", pOKMsg, If(aErr = False, pOKMsg & Create, "Error" & Create))
        Catch Ex As System.Exception
            MsgBox("Error: Exception " & Ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAPRouting")
            Create = "Error: Exception in Change"
        Finally
            RfcSessionManager.EndContext(destination)
        End Try
    End Function

    Public Function Change(pData As TSAP_RoutingData, Optional pOKMsg As String = "OK", Optional pCheck As Boolean = False) As String
        Change = ""
        Try
            oRfcFunction = destination.Repository.CreateFunction("ROUTING_MAINTAIN")
            RfcSessionManager.BeginContext(destination)
            Dim oRETURN As IRfcTable = oRfcFunction.GetTable("RETURN")
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
            For Each aKey In pData.aDataDic.aTDataDic.Keys
                Dim oTable As IRfcTable = oRfcFunction.GetTable(aKey)
                oTable.Clear()
                pData.aDataDic.to_IRfcTable(pKey:=aKey, pIRfcTable:=oTable)
            Next
            ' call the BAPI
            oRfcFunction.Invoke(destination)
            Dim aErr As Boolean = False
            For i As Integer = 0 To oRETURN.Count - 1
                Change = Change & ";" & oRETURN(i).GetValue("MESSAGE")
                If oRETURN(i).GetValue("TYPE") <> "S" And oRETURN(i).GetValue("TYPE") <> "I" And oRETURN(i).GetValue("TYPE") <> "W" Then
                    aErr = True
                End If
            Next
            If aErr = False Then
                Dim aSAPBapiTranctionCommit As New SAPBapiTranctionCommit(sapcon)
                aSAPBapiTranctionCommit.commit(pWait:="X")
            End If
            Change = If(Change = "", pOKMsg, If(aErr = False, pOKMsg & Change, "Error" & Change))
        Catch Ex As System.Exception
            MsgBox("Error: Exception " & Ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAPRouting")
            Change = "Error: Exception in Change"
        Finally
            RfcSessionManager.EndContext(destination)
        End Try
    End Function

End Class
