Imports SAP.Middleware.Connector

Public Class TDataDic

    Public aTDataDic As Dictionary(Of String, TData)
    Private aPar As SAPCommon.TStr

    Public Sub New(ByRef pPar As SAPCommon.TStr)
        aTDataDic = New Dictionary(Of String, TData)
        aPar = pPar
    End Sub

    Public Sub addValue(pStruc As String, pKey As String, pNAME As String, pVALUE As String, pCURRENCY As String, pFORMAT As String,
                        Optional pEmty As Boolean = False, Optional pEmptyChar As String = "#", Optional pOperation As String = "set")
        Dim aTData As TData

        If aTDataDic.ContainsKey(pStruc) Then
            aTData = aTDataDic(pStruc)
            aTData.addValue(pKey, pNAME, pVALUE, pCURRENCY, pFORMAT, pEmty, pEmptyChar, pOperation)
        Else
            aTData = New TData(aPar)
            aTData.addValue(pKey, pNAME, pVALUE, pCURRENCY, pFORMAT, pEmty, pEmptyChar, pOperation)
            aTDataDic.Add(pStruc, aTData)
        End If
    End Sub

    Public Sub addValue(pKey As String, ByRef oStruc As IRfcStructure, Optional pStrucName As String = "")
        If Not oStruc Is Nothing Then
            Dim aStrucName As String = If(pStrucName = "", oStruc.Metadata.Name, pStrucName)
            For j As Integer = 0 To oStruc.Count - 1
                addValue(aStrucName, pKey, aStrucName & "-" & oStruc(j).Metadata.Name, CStr(oStruc(j).GetValue), "", "")
            Next
        End If
    End Sub

    Public Sub addValues(ByRef oTable As IRfcTable, Optional pStrucName As String = "")
        Dim aStrucName As String
        If Not oTable Is Nothing Then
            aStrucName = If(pStrucName = "", oTable(0).Metadata.Name, pStrucName)
            For i As Integer = 0 To oTable.Count - 1
                addValue(CStr(i), oTable(i), aStrucName)
            Next
        End If
    End Sub

    Public Sub addValue(pKey As String, pTStrRec As SAPCommon.TStrRec,
                        Optional pEmty As Boolean = False, Optional pEmptyChar As String = "#", Optional pOperation As String = "set",
                        Optional pNewStrucname As String = "")
        Dim aTData As TData
        Dim aName As String
        Dim aStruc As String
        If pNewStrucname <> "" Then
            aStruc = pNewStrucname
            aName = pNewStrucname & "-" & pTStrRec.Fieldname
        Else
            aStruc = pTStrRec.Strucname
            aName = pTStrRec.Strucname & "-" & pTStrRec.Fieldname
        End If

        If aTDataDic.ContainsKey(aStruc) Then
            aTData = aTDataDic(aStruc)
            aTData.addValue(pKey, aName, pTStrRec.Value, pTStrRec.Currency, pTStrRec.Format, pEmty, pEmptyChar, pOperation)
        Else
            aTData = New TData(aPar)
            aTData.addValue(pKey, aName, pTStrRec.Value, pTStrRec.Currency, pTStrRec.Format, pEmty, pEmptyChar, pOperation)
            aTDataDic.Add(aStruc, aTData)
        End If
    End Sub

    Public Sub to_IRfcTable(pKey As String, ByRef pIRfcTable As IRfcTable)
        Dim aTData As TData
        Dim aKvP As KeyValuePair(Of String, TDataRec)
        Dim aTDataRec As TDataRec
        If aTDataDic.ContainsKey(pKey) Then
            aTData = aTDataDic(pKey)
            For Each aKvP In aTData.aTDataDic
                Dim oAppended As Boolean = False
                aTDataRec = aKvP.Value
                For Each aTStrRec In aTDataRec.aTDataRecCol
                    If Not oAppended Then
                        pIRfcTable.Append()
                        oAppended = True
                    End If
                    pIRfcTable.SetValue(CType(aTStrRec.Fieldname, String), CType(aTStrRec.formated(), String))
                Next
            Next
        End If
    End Sub

    Public Sub to_IRfcTableX(pKey As String, ByRef pIRfcTable As IRfcTable, Optional pKeyFields() As String = Nothing)
        Dim aTData As TData
        Dim aKvP As KeyValuePair(Of String, TDataRec)
        Dim aTDataRec As TDataRec
        If aTDataDic.ContainsKey(pKey) Then
            aTData = aTDataDic(pKey)
            For Each aKvP In aTData.aTDataDic
                Dim oAppended As Boolean = False
                aTDataRec = aKvP.Value
                For Each aTStrRec In aTDataRec.aTDataRecCol
                    If Not oAppended Then
                        pIRfcTable.Append()
                        oAppended = True
                    End If
                    If Not pKeyFields Is Nothing Then
                        If isInArray(aTStrRec.Fieldname, pKeyFields) Then
                            pIRfcTable.SetValue(CType(aTStrRec.Fieldname, String), CType(aTStrRec.formated(), String))
                        Else
                            pIRfcTable.SetValue(CType(aTStrRec.Fieldname, String), "X")
                        End If
                    Else
                        pIRfcTable.SetValue(CType(aTStrRec.Fieldname, String), "X")
                    End If
                Next
            Next
        End If
    End Sub

    Private Function isInArray(pString As String, pArray As Object) As Boolean
        Dim st As String, M As String
        M = "$"
        st = M & Join(pArray, M) & M
        isInArray = InStr(st, M & pString & M) > 0
    End Function

End Class
