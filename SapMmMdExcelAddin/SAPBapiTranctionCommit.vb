' Copyright 2016-2019 Hermann Mundprecht
' This file is licensed under the terms of the license 'CC BY 4.0'. 
' For a human readable version of the license, see https://creativecommons.org/licenses/by/4.0/

Imports SAP.Middleware.Connector

Public Class SAPBapiTranctionCommit

    Private Shared ReadOnly log As log4net.ILog = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType)
    Private oRfcFunction As IRfcFunction
    Private destination As RfcCustomDestination
    Private sapcon As SapCon

    Sub New(aSapCon As SapCon)
        sapcon = aSapCon
        aSapCon.getDestination(destination)
        log.Debug("New - " & "creating Function BAPI_TRANSACTION_COMMIT")
        Try
            oRfcFunction = destination.Repository.CreateFunction("BAPI_TRANSACTION_COMMIT")
            log.Debug("New - " & "oRfcFunction.Metadata.Name=" & oRfcFunction.Metadata.Name)
        Catch Exc As System.Exception
            log.Error("New - Exception=" & Exc.ToString)
        End Try
    End Sub

    Public Function commit(Optional pWait As String = "") As Integer
        sapcon.checkCon()
        Try
            log.Debug("commit - " & "invoking " & oRfcFunction.Metadata.Name & "pWait=" & pWait)
            oRfcFunction.SetValue("WAIT", pWait)
            oRfcFunction.Invoke(destination)
            commit = 0
            Exit Function
        Catch ex As Exception
            log.Error("commit - Exception=" & ex.ToString)
            MsgBox("Exception in commit! " & ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAPBapiTranctionCommit")
            commit = 8
        End Try

    End Function
End Class
