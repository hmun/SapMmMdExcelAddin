' Copyright 2020 Hermann Mundprecht
' This file is licensed under the terms of the license 'CC BY 4.0'. 
' For a human readable version of the license, see https://creativecommons.org/licenses/by/4.0/

Public Class TDataRec

    Public aTDataRecCol As Collection
    Private aIntPar As SAPCommon.TStr

    Public Sub New(ByRef pIntPar As SAPCommon.TStr)
        aTDataRecCol = New Collection
        aIntPar = pIntPar
    End Sub

    Public Sub setValues(pNAME As String, pVALUE As String, pCURRENCY As String, pFORMAT As String,
                         Optional pEmty As Boolean = False, Optional pEmptyChar As String = "#", Optional pOperation As String = "set", Optional pUseAsEmpty As String = "#")
        Dim aTStrRec As SAPCommon.TStrRec
        Dim aNameArray() As String
        Dim aKey As String
        Dim aSTRUCNAME As String = ""
        Dim aFIELDNAME As String = ""
        Dim aValue As String
        If pVALUE = pUseAsEmpty Then
            aValue = " "
        Else
            aValue = pVALUE
            If Not pEmty And aValue = pEmptyChar Then
                Exit Sub
            End If
        End If
        ' do not add empty values

        If InStr(pNAME, "-") <> 0 Then
            aNameArray = Split(pNAME, "-")
            aSTRUCNAME = aNameArray(0)
            aFIELDNAME = aNameArray(1)
        Else
            aSTRUCNAME = ""
            aFIELDNAME = pNAME
        End If
        aKey = pNAME
        If aTDataRecCol.Contains(aKey) Then
            aTStrRec = aTDataRecCol(aKey)
            Select Case pOperation
                Case "add"
                    aTStrRec.addValues(aSTRUCNAME, aFIELDNAME, aValue, pCURRENCY, pFORMAT)
                Case "sub"
                    aTStrRec.subValues(aSTRUCNAME, aFIELDNAME, aValue, pCURRENCY, pFORMAT)
                Case "mul"
                    aTStrRec.mulValues(aSTRUCNAME, aFIELDNAME, aValue, pCURRENCY, pFORMAT)
                Case "div"
                    aTStrRec.divValues(aSTRUCNAME, aFIELDNAME, aValue, pCURRENCY, pFORMAT)
                Case Else
                    aTStrRec.setValues(aSTRUCNAME, aFIELDNAME, aValue, pCURRENCY, pFORMAT)
            End Select
        Else
            aTStrRec = New SAPCommon.TStrRec
            aTStrRec.setValues(aSTRUCNAME, aFIELDNAME, aValue, pCURRENCY, pFORMAT)
            aTDataRecCol.Add(aTStrRec, aKey)
        End If
    End Sub

    Public Sub setValues(pTDataRec As TDataRec, Optional pEmty As Boolean = False, Optional pEmptyChar As String = "#", Optional pOperation As String = "set")
        Dim aTStrRec As SAPCommon.TStrRec
        For Each aTStrRec In pTDataRec.aTDataRecCol
            setValues(aTStrRec.getKey(), aTStrRec.Value, aTStrRec.Currency, aTStrRec.Format, pEmty, pEmptyChar, pOperation)
        Next
    End Sub

    Public Sub addValues(pTDataRec As TDataRec, Optional pEmty As Boolean = False, Optional pEmptyChar As String = "#")
        Dim aTStrRec As SAPCommon.TStrRec
        For Each aTStrRec In pTDataRec.aTDataRecCol
            If aTStrRec.Currency <> "" Then
                setValues(aTStrRec.getKey(), aTStrRec.Value, aTStrRec.Currency, aTStrRec.Format, pEmty, pEmptyChar, pOperation:="add")
            Else
                setValues(aTStrRec.getKey(), aTStrRec.Value, aTStrRec.Currency, aTStrRec.Format, pEmty, pEmptyChar, pOperation:="set")
            End If
        Next
    End Sub

    Public Function getColumn(pClmn As String) As SAPCommon.TStrRec
        Dim aTStrRec As SAPCommon.TStrRec
        If aTDataRecCol.Contains(pClmn) Then
            aTStrRec = aTDataRecCol(pClmn)
            getColumn = aTStrRec
        Else
            getColumn = Nothing
        End If
    End Function

    Public Function getMaterial() As SAPCommon.TStrRec
        Dim aTlClmn As String = If(aIntPar.value("COL", "MATERIAL") <> "", aIntPar.value("COL", "MATERIAL"), "MATERIAL")
        getMaterial = getColumn(aTlClmn)
    End Function

    Public Function getMaintStat() As SAPCommon.TStrRec
        getMaintStat = getColumn("CLIENTDATA-MAINT_STAT")
    End Function

    Public Function getMvgPrice() As SAPCommon.TStrRec
        Dim aTStrRec As SAPCommon.TStrRec
        getMvgPrice = getColumn("VALUATIONDATA-MOVING_PR")

    End Function

    Public Function getStdPrice() As SAPCommon.TStrRec
        getStdPrice = getColumn("VALUATIONDATA-STD_PRICE")
    End Function

    Public Function getPriceUnit() As SAPCommon.TStrRec
        getPriceUnit = getColumn("VALUATIONDATA-PRICE_UNIT")
    End Function

    Public Function getValArea() As SAPCommon.TStrRec
        Dim aTStrRec As SAPCommon.TStrRec
        getValArea = getColumn("VALUATIONDATA-VAL_AREA")
    End Function

    Public Function getValType() As SAPCommon.TStrRec
        Dim aTStrRec As SAPCommon.TStrRec
        getValType = getColumn("VALUATIONDATA-VAL_TYPE")
    End Function

    Public Function getValidFrom() As SAPCommon.TStrRec
        Dim aTStrRec As SAPCommon.TStrRec
        getValidFrom = getColumn("VALUATIONDATA-VALID_FROM")
    End Function

    Public Function getPost(ByRef pPar As SAPCommon.TStr) As String
        Dim aClmn As String = If(pPar.value("COL", "DATAPOST") <> "", pPar.value("COL", "DATAPOST"), "INT-POST")
        Dim aTStrRec As SAPCommon.TStrRec = getColumn(aClmn)
        getPost = If(aTStrRec Is Nothing, "", aTStrRec.Value)
    End Function

    Public Function hasCurPrice() As Boolean
        Dim aStrucName() As String
        hasCurPrice = False
        Dim aTStrRec As SAPCommon.TStrRec
        For Each aTStrRec In aTDataRecCol
            aStrucName = Split(aTStrRec.Strucname, "+")
            If (isInArray("A00", aStrucName) Or isInArray("A10", aStrucName) Or isInArray("A20", aStrucName) Or isInArray("A30", aStrucName) Or isInArray("A40", aStrucName)) And aTStrRec.Fieldname = "PRICE" Then
                hasCurPrice = True
                Exit Function
            End If
        Next
    End Function

    Public Sub toRange(pFields() As String, pIsValue() As String, ByRef aRange As Excel.Range)
        Dim aTStrRec As SAPCommon.TStrRec
        For i = 0 To pFields.Count - 1
            If aTDataRecCol.Contains(pFields(i)) Then
                aTStrRec = aTDataRecCol(pFields(i))
                If pIsValue(i) = "X" Then
                    aRange(1, i + 1).Value = CDbl(aTStrRec.formated())
                Else
                    aRange(1, i + 1).Value = CStr(aTStrRec.formated())
                End If
            End If
        Next
    End Sub

    Private Function isInArray(pString As String, pArray As Object) As Boolean
        Dim st As String, M As String
        M = "$"
        st = M & Join(pArray, M) & M
        isInArray = InStr(st, M & pString & M) > 0
    End Function

End Class
