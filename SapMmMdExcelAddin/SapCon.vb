' Copyright 2016-2019 Hermann Mundprecht
' This file is licensed under the terms of the license 'CC BY 4.0'. 
' For a human readable version of the license, see https://creativecommons.org/licenses/by/4.0/

Imports SAP.Middleware.Connector
Public Class SapCon
    Const aParamWs As String = "Parameter"
    Const aConnectionWs As String = "SAP-Con"
    Private Shared ReadOnly log As log4net.ILog = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType)
    Private _sapcon As SAPLogon.SAPCon

    Public Sub New()
        Dim conParameter As New SAPLogon.ConParameter
        Dim aSapExcelDestination As New SapExcelDestination
        Dim aCws As Excel.Worksheet
        Dim aWB As Excel.Workbook
        aWB = Globals.SapMmMdExcelAddin.Application.ActiveWorkbook
        Dim assemblyName As System.Reflection.AssemblyName = System.Reflection.Assembly.GetExecutingAssembly().GetName()
        Dim assembly As String = assemblyName.Name
        Try
            aCws = aWB.Worksheets(aConnectionWs)
        Catch Exc As System.Exception
            MsgBox("No " & aConnectionWs & " Sheet in current workbook", MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SapCon")
            log.Error("New - Exception=" & Exc.ToString)
            Exit Sub
        End Try
        log.Debug("New - " & "setting up aSapExcelDestinationConfiguration")
        conParameter = aSapExcelDestination.GetExcelDestinations(aConnectionWs)
        _sapcon = New SAPLogon.SAPCon(assembly, conParameter)
        '        _sapcon.Dest = CStr(aCws.Cells(2, 2).Value)
    End Sub

    Public Function getDestination(ByRef pDestination As RfcCustomDestination) As Boolean
        getDestination = False
        pDestination = _sapcon.Destination
        If Not pDestination Is Nothing Then
            getDestination = True
        End If
    End Function

    Public Function getSapCon(ByRef pSapCon As SAPLogon.SAPCon) As Boolean
        getSapCon = False
        pSapCon = _sapcon
        If Not pSapCon Is Nothing Then
            getSapCon = True
        End If
    End Function

    Private Sub setDest()
        Dim formRet = 0
        Dim oForm As New FormDestinations
        Dim destCol As System.Collections.ObjectModel.Collection(Of String)
        Dim dest As String
        log.Debug("setDest - " & "building destination list")
        destCol = _sapcon.GetDestinationList()
        For Each dest In destCol
            oForm.ListBoxDest.Items.Add(dest)
        Next
        formRet = oForm.ShowDialog()
        If formRet = System.Windows.Forms.DialogResult.OK Then
            _sapcon.Dest = oForm.ListBoxDest.SelectedItem.ToString
            log.Debug("setDest - " & "selected aDest=" & _sapcon.Dest)
        Else
            log.Debug("setDest - " & "no destination selected")
            _sapcon.Dest = ""
        End If
    End Sub

    Public Function checkCon() As Integer
        Dim formRet = 0
        Dim aRet As Integer
        If _sapcon.Dest = "" Then
            setDest()
        End If
        If _sapcon.Destination Is Nothing Then
            aRet = _sapcon.setDestination()
            If aRet <> 0 Then
                MsgBox("Error reading destination " & _sapcon.Dest & "! Check the connection settings in the sap_connections.config file and the SAP-Con sheet",
                        MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SapCon")
                checkCon = aRet
            End If
        End If

        If Not _sapcon.Connected And CStr(_sapcon.Destination.SncMode) = "1" Then
            Dim oForm As New FormLogon
            Dim aClient As String = ""
            Dim aUserName As String = ""
            Dim aSncMyName As String = ""
            Dim aPassword As String = ""
            Dim aLanguage As String = ""
            log.Debug("checkCon - " & "connecting using SNC destination")
            oForm.isSNC = True
            oForm.Destination.Text = _sapcon.Dest
            If Not _sapcon.Destination.Client Is Nothing Then
                oForm.Client.Text = _sapcon.Destination.Client
            End If
            If My.Settings.SAP_Language IsNot Nothing And My.Settings.SAP_Language <> "" Then
                oForm.Language.Text = My.Settings.SAP_Language
            ElseIf Not _sapcon.Destination.Language Is Nothing Then
                oForm.Language.Text = _sapcon.Destination.Language
            End If
            If Not _sapcon.Destination.SncMyName Is Nothing Then
                oForm.SNCName.Text = _sapcon.Destination.SncMyName
            ElseIf My.Settings.SAP_SncMyName IsNot Nothing And My.Settings.SAP_SncMyName <> "" Then
                oForm.SNCName.Text = My.Settings.SAP_SncMyName
            End If
            oForm.UserName.Text = ""
            oForm.UserName.Enabled = True
            oForm.Password.Enabled = False
            formRet = oForm.ShowDialog()
            If formRet = System.Windows.Forms.DialogResult.OK Then
                aClient = oForm.Client.Text
                If Not String.IsNullOrEmpty(oForm.UserName.Text) Then
                    aUserName = oForm.UserName.Text
                End If
                aPassword = oForm.Password.Text
                aLanguage = oForm.Language.Text
                If Not String.IsNullOrEmpty(oForm.SNCName.Text) Then
                    aSncMyName = oForm.SNCName.Text
                End If
                My.Settings.SAP_Language = oForm.Language.Text
                _sapcon.Client = aClient
                _sapcon.Language = aLanguage
                _sapcon.SncMyName = aSncMyName
                _sapcon.Username = aUserName
            End If
        ElseIf Not _sapcon.Connected Then
            Dim oForm As New FormLogon
            Dim aClient As String
            Dim aUserName As String
            Dim aPassword As String
            Dim aLanguage As String
            log.Debug("checkCon - " & "connecting using regular destination")
            If Not _sapcon.Destination.Client Is Nothing Then
                oForm.Client.Text = _sapcon.Destination.Client
            End If
            If My.Settings.SAP_Language IsNot Nothing And My.Settings.SAP_Language <> "" Then
                oForm.Language.Text = My.Settings.SAP_Language
            ElseIf Not _sapcon.Destination.Language Is Nothing Then
                oForm.Language.Text = _sapcon.Destination.Language
            End If
            oForm.isSNC = False
            oForm.Destination.Text = _sapcon.Dest
            oForm.UserName.Enabled = True
            If My.Settings.SAP_User IsNot Nothing Then
                oForm.UserName.Text = CStr(My.Settings.SAP_User)
            End If
            oForm.Password.Enabled = True
            oForm.SNCName.Enabled = False
            formRet = oForm.ShowDialog()
            If formRet = System.Windows.Forms.DialogResult.OK Then
                aClient = oForm.Client.Text
                aUserName = oForm.UserName.Text
                My.Settings.SAP_User = oForm.UserName.Text
                aPassword = oForm.Password.Text
                aLanguage = oForm.Language.Text
                My.Settings.SAP_Language = oForm.Language.Text
                _sapcon.Client = aClient
                _sapcon.Username = aUserName
                _sapcon.Password = aPassword
                _sapcon.Language = aLanguage
            End If
        End If
        Try
            If formRet = System.Windows.Forms.DialogResult.OK Then
                aRet = _sapcon.checkCon(True)
            Else
                aRet = _sapcon.checkCon(False)
            End If
            checkCon = aRet
        Catch ex As RfcInvalidParameterException
            MsgBox("Connecting to SAP failed! " & ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SapCon")
            checkCon = 4
        Catch ex As RfcBaseException
            MsgBox("Connecting to SAP failed! " & ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SapCon")
            checkCon = 8
        End Try
    End Function

    Public Sub SAPlogoff()
        _sapcon.SAPlogoff()
    End Sub

End Class
