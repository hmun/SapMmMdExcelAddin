﻿' Copyright 2022 Hermann Mundprecht
' This file is licensed under the terms of the license 'CC BY 4.0'. 
' For a human readable version of the license, see https://creativecommons.org/licenses/by/4.0/

Imports System.Collections.Specialized
Imports System.Configuration
Imports System.Net.PeerToPeer.Collaboration
Imports Microsoft.Office.Tools.Ribbon

Public Class SapMmMdRibbon

    Private aSapCon
    Private aSapGeneral
    Private aTlPar As SAPCommon.TStr
    Private aIntPar As SAPCommon.TStr
    Private Shared ReadOnly log As log4net.ILog = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType)

    Private Sub SapMmMdRibbon_Load(ByVal sender As System.Object, ByVal e As RibbonUIEventArgs) Handles MyBase.Load
        Dim enableSourceList As Boolean = False
        Dim enableRouting As Boolean = False
        Dim enablePriceChange As Boolean = False
        Dim enableGoodsMovement As Boolean = False
        aSapGeneral = New SapGeneral
        Try
            enableSourceList = Convert.ToBoolean(ConfigurationManager.AppSettings("enableSourceList"))
            enableRouting = Convert.ToBoolean(ConfigurationManager.AppSettings("enableRouting"))
            enablePriceChange = Convert.ToBoolean(ConfigurationManager.AppSettings("enablePriceChange"))
            enableGoodsMovement = Convert.ToBoolean(ConfigurationManager.AppSettings("enableGoodsMovement"))
        Catch Exc As System.Exception
            log.Error("SapAccRibbon_Load - " & "Exception=" & Exc.ToString)
        End Try
        Globals.Ribbons.Ribbon1.GroupSourceList.Visible = enableSourceList
        Globals.Ribbons.Ribbon1.GroupRouting.Visible = enableRouting
        Globals.Ribbons.Ribbon1.GroupMaterialPrice.Visible = enablePriceChange
        Globals.Ribbons.Ribbon1.GroupGoodsMovement.Visible = enableGoodsMovement
    End Sub

    Private Function checkCon() As Integer
        Dim aSapConRet As Integer
        Dim aSapVersionRet As Integer
        checkCon = False
        log.Debug("checkCon - " & "checking Version")
        If Not aSapGeneral.checkVersion() Then
            Exit Function
        End If
        log.Debug("checkCon - " & "checking Connection")
        aSapConRet = 0
        If aSapCon Is Nothing Then
            Try
                aSapCon = New SapCon()
            Catch ex As SystemException
                log.Warn("checkCon-New SapCon - )" & ex.ToString)
            End Try
        End If
        Try
            aSapConRet = aSapCon.checkCon()
        Catch ex As SystemException
            log.Warn("checkCon-aSapCon.checkCon - )" & ex.ToString)
        End Try
        If aSapConRet = 0 Then
            log.Debug("checkCon - " & "checking version in SAP")
            Try
                aSapVersionRet = aSapGeneral.checkVersionInSAP(aSapCon)
            Catch ex As SystemException
                log.Warn("checkCon - )" & ex.ToString)
            End Try
            log.Debug("checkCon - " & "aSapVersionRet=" & CStr(aSapVersionRet))
            If aSapVersionRet = True Then
                log.Debug("checkCon - " & "checkCon = True")
                checkCon = True
            Else
                log.Debug("checkCon - " & "connection check failed")
            End If
        End If
    End Function

    Private Sub ButtonLogoff_Click(sender As Object, e As RibbonControlEventArgs) Handles ButtonLogoff.Click
        log.Debug("ButtonLogoff_Click - " & "starting logoff")
        If Not aSapCon Is Nothing Then
            log.Debug("ButtonLogoff_Click - " & "calling aSapCon.SAPlogoff()")
            aSapCon.SAPlogoff()
            aSapCon = Nothing
        End If
        log.Debug("ButtonLogoff_Click - " & "exit")
    End Sub

    Private Sub ButtonLogon_Click(sender As Object, e As RibbonControlEventArgs) Handles ButtonLogon.Click
        Dim aConRet As Integer

        log.Debug("ButtonLogon_Click - " & "checking Version")
        If Not aSapGeneral.checkVersion() Then
            log.Debug("ButtonLogon_Click - " & "Version check failed")
            Exit Sub
        End If
        log.Debug("ButtonLogon_Click - " & "creating SapCon")
        If aSapCon Is Nothing Then
            aSapCon = New SapCon()
        End If
        log.Debug("ButtonLogon_Click - " & "calling SapCon.checkCon()")
        aConRet = aSapCon.checkCon()
        If aConRet = 0 Then
            log.Debug("ButtonLogon_Click - " & "connection successfull")
            MsgBox("SAP-Logon successful! ", MsgBoxStyle.OkOnly Or MsgBoxStyle.Information, "Sap LTP")
        Else
            log.Debug("ButtonLogon_Click - " & "connection failed")
            aSapCon = Nothing
        End If
    End Sub

    Private Sub ButtonSapMaterialChange_Click(sender As Object, e As RibbonControlEventArgs) Handles ButtonSapMaterialChange.Click
        Dim aSapMmMdRibbon_Mat As New SapMmMdRibbon_Mat
        If checkCon() = True Then
            aSapMmMdRibbon_Mat.Change(pSapCon:=aSapCon)
        Else
            MsgBox("Checking SAP-Connection failed!", MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap ButtonSapMaterialChange_Click")
        End If
    End Sub

    Private Sub ButtonSapMaterialGetAll_Click(sender As Object, e As RibbonControlEventArgs) Handles ButtonSapMaterialGetAll.Click
        Dim aSapMmMdRibbon_Mat As New SapMmMdRibbon_Mat
        If checkCon() = True Then
            aSapMmMdRibbon_Mat.GetAll(pSapCon:=aSapCon)
        Else
            MsgBox("Checking SAP-Connection failed!", MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap ButtonSapMaterialChange_Click")
        End If
    End Sub

    Private Sub ButtonSapMaterialPriceChange_Click(sender As Object, e As RibbonControlEventArgs) Handles ButtonSapMaterialPriceChange.Click
        Dim aSapMmMdRibbon_Mat As New SapMmMdRibbon_Mat
        If checkCon() = True Then
            aSapMmMdRibbon_Mat.PriceChange(pSapCon:=aSapCon)
        Else
            MsgBox("Checking SAP-Connection failed!", MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap ButtonSapMaterialPriceChange_Click")
        End If
    End Sub

    Private Sub ButtonSapSourceListRead_Click(sender As Object, e As RibbonControlEventArgs) Handles ButtonSapSourceListRead.Click
        Dim aSapMmMdRibbon_SL As New SapMmMdRibbon_SL
        If checkCon() = True Then
            aSapMmMdRibbon_SL.Read(pSapCon:=aSapCon)
        Else
            MsgBox("Checking SAP-Connection failed!", MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap ButtonSapSourceListRead_Click")
        End If
    End Sub

    Private Sub ButtonSapSourceListUpdate_Click(sender As Object, e As RibbonControlEventArgs) Handles ButtonSapSourceListUpdate.Click
        Dim aSapMmMdRibbon_SL As New SapMmMdRibbon_SL
        If checkCon() = True Then
            aSapMmMdRibbon_SL.Update(pSapCon:=aSapCon)
        Else
            MsgBox("Checking SAP-Connection failed!", MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap ButtonSapSourceListUpdate_Click")
        End If
    End Sub

    Private Sub ButtonSapRoutingCreate_Click(sender As Object, e As RibbonControlEventArgs) Handles ButtonSapRoutingCreate.Click
        Dim aSapMmMdRibbon_Routing As New SapMmMdRibbon_Routing
        If checkCon() = True Then
            aSapMmMdRibbon_Routing.Maintain(pSapCon:=aSapCon, pMode:="create")
        Else
            MsgBox("Checking SAP-Connection failed!", MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap ButtonSapRoutingCreate_Click")
        End If
    End Sub

    Private Sub ButtonSapRoutingChange_Click(sender As Object, e As RibbonControlEventArgs) Handles ButtonSapRoutingChange.Click
        Dim aSapMmMdRibbon_Routing As New SapMmMdRibbon_Routing
        If checkCon() = True Then
            aSapMmMdRibbon_Routing.Maintain(pSapCon:=aSapCon, pMode:="change")
        Else
            MsgBox("Checking SAP-Connection failed!", MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap ButtonSapRoutingChange_Click")
        End If
    End Sub

    Private Sub ButtonSapGoodsMovementCreate_Click(sender As Object, e As RibbonControlEventArgs) Handles ButtonSapGoodsMovementCreate.Click
        Dim aSapMmMdRibbon_GoodsMovement As New SapMmMdRibbon_GoodsMovement
        If checkCon() = True Then
            aSapMmMdRibbon_GoodsMovement.Maintain(pSapCon:=aSapCon, pMode:="create")
        Else
            MsgBox("Checking SAP-Connection failed!", MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap ButtonSapGoodsMovementCreate_Click")
        End If
    End Sub
End Class
