'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€
' Guanzon Software Engineering Group
' Guanzon Group of Companies
' Perez Blvd., Dagupan City
'
'   GOCAS Number Decoder
'
' Copyright 2012 and Beyond
' All Rights Reserved
' ºººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººº
' €  All  rights reserved. No part of this  software  €€  This Software is Owned by        €
' €  may be reproduced or transmitted in any form or  €€                                   €
' €  by   any   means,  electronic   or  mechanical,  €€    GUANZON MERCHANDISING CORP.    €
' €  including recording, or by information  storage  €€     Guanzon Bldg. Perez Blvd.     €
' €  and  retrieval  systems, without  prior written  €€           Dagupan City            €
' €  from the author.                                 €€  Tel No. 522-1085 ; 522-9275      €
' ºººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººº
'   Logic by Sir Marlon Sayson
' ==========================================================================================
'   Mac 2019.12.11 08:30 AM
'       Started creating this object.
'   Mac 2020.06.10 04:00 pM
'       replace special chars like period to empty if the name was not empty, if empty set it to zero
' ==========================================================================================
'   IMPORTANT FIELDS:
'       sTransNox - Credit Online Application Transaction No.
'       sUserIDxx - Account ID of the App User
'       sClientNm - Customer Name (LastName, FirstName SuffixName Middle Name)
'       nDownPaym - Integer value of rounded required downpayment percentage
'       cWithCIxx - Is credit investigation required? 0 or 1
'       sGOCASNox - GOCAS Number Result
' ==========================================================================================
'  PLEASE SEE modCrypt.Main() for test USAGE.
' ==========================================================================================
'  DP = «w/CI»(N) + LPAD(«DP», 3, "")(NNN)
'  DP = NNNN + random1 + random2
'
'  DP = 1200 - default downpayment and requires CI
'  DP = 200 - default downpayment and requires NO CI
'  DP = 1100 - disapproved(cash option only)
'  DP = 100 - disapproved(cash option only)
'  DP = 1010 - 10% DP and requires CI
'  DP = 10 - 10% DP and requires NO CI
'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€
Option Explicit On

Public Class GOCASCodeGen
    Private Const BRANCHCD As String = "GAP0"

    Dim sTransNox As String
    Dim sUserIDxx As String
    Dim sClientNm As String
    Dim cWithCIxx As String
    Dim sGOCASNox As String

    'applicant name
    Dim sLastName As String
    Dim sFrstName As String
    Dim sMiddName As String
    Dim sSuffixNm As String

    Dim nDownPaym As Integer

    Dim sWarnMsgx As String

    Property TransactionNo() As String
        Get
            Return sTransNox
        End Get
        Set(ByVal Value As String)
            sTransNox = Value
        End Set
    End Property

    Property UserID() As String
        Get
            Return sUserIDxx
        End Get
        Set(ByVal Value As String)
            sUserIDxx = Value
        End Set
    End Property

    Property LastName() As String
        Get
            Return sLastName
        End Get
        Set(ByVal Value As String)
            If Len(Value) = 1 Then
                sLastName = Trim(Replace(Replace(Replace(Replace(Value, ".", "0"), "'", ""), ",", "0"), "-", "0"))
            Else
                sLastName = Trim(Replace(Replace(Replace(Replace(Value, ".", ""), "'", ""), ",", ""), "-", ""))
            End If            
        End Set
    End Property

    Property FirstName() As String
        Get
            Return sFrstName
        End Get
        Set(ByVal Value As String)
            If Len(Value) = 1 Then
                sFrstName = Trim(Replace(Replace(Replace(Replace(Value, ".", "0"), "'", ""), ",", "0"), "-", "0"))
            Else
                sFrstName = Trim(Replace(Replace(Replace(Replace(Value, ".", ""), "'", ""), ",", ""), "-", ""))
            End If
        End Set
    End Property

    Property MiddleName() As String
        Get
            Return sMiddName
        End Get
        Set(ByVal Value As String)
            Value = Trim(Value)

            If Len(Value) = 1 Then
                sMiddName = Trim(Replace(Replace(Replace(Replace(Value, ".", "0"), "'", ""), ",", "0"), "-", "0"))
            Else
                sMiddName = Trim(Replace(Replace(Replace(Replace(Value, ".", ""), "'", ""), ",", ""), "-", ""))
            End If
        End Set
    End Property

    Property SuffixName() As String
        Get
            Return sSuffixNm
        End Get
        Set(ByVal Value As String)
            If Len(Value) = 1 Then
                sSuffixNm = Trim(Replace(Replace(Replace(Replace(Value, ".", "0"), "'", ""), ",", "0"), "-", "0"))
            Else
                sSuffixNm = Trim(Replace(Replace(Replace(Replace(Value, ".", ""), "'", ""), ",", ""), "-", ""))
            End If
        End Set
    End Property

    Property IsCINeeded() As Boolean
        Get
            Return CBool(IIf(cWithCIxx = "1", True, False))
        End Get
        Set(ByVal Value As Boolean)
            cWithCIxx = CStr(IIf(Value = True, "1", "0"))
        End Set
    End Property

    Property DownPayment() As Integer
        Get
            Return nDownPaym
        End Get
        Set(ByVal Value As Integer)
            nDownPaym = Value
        End Set
    End Property


    ReadOnly Property GOCASApprvl() As String
        Get
            Return sGOCASNox
        End Get
    End Property

    Property Message() As String
        Get
            Return sWarnMsgx
        End Get
        Set(ByVal Value As String)
            sWarnMsgx = Value
        End Set
    End Property

    Public Function IsValidGOCASNo(ByVal fsTransNox As String, _
                                   ByVal fsGOCASNox As String, _
                                   ByVal fsLastName As String, _
                                   ByVal fsFrstName As String, _
                                   ByVal fsMiddName As String, _
                                   ByVal fsSuffixNm As String, _
                                   ByVal fnUnitPrce As Double, _
                                   ByVal fnDownPaym As Double) As Boolean

        Message = "Invalid GOCAS Number Detected."

        If (Not Decode(fsGOCASNox)) Then Return False

        If sTransNox <> fsTransNox Then Return False

        'get the client code
        Dim lsName As String = ClientCode(fsLastName, fsFrstName, fsMiddName, fsSuffixNm)

        'compare the decoded client code to the client code from the given client name
        If sClientNm <> lsName Then Return False

        Dim lnDownPaym As Double = (fnDownPaym / fnUnitPrce) * 100

        If nDownPaym > lnDownPaym Then Return False

        Return True
    End Function

    Public Function IsValidGOCASNo(ByVal fsGOCASNox As String, _
                                   ByVal fsClientNm As String, _
                                   ByVal fnUnitPrce As Double, _
                                   ByVal fnDownPaym As Double) As Boolean

        Message = "Invalid GOCAS Number Detected."

        If (Not Decode(fsGOCASNox)) Then Return False

        'get the client code
        Dim lsName As String = ClientCode(fsClientNm)

        'compare the decoded client code to the client code from the given client name
        If sClientNm <> lsName Then Return False

        Dim lnDownPaym As Double = (fnDownPaym / fnUnitPrce) * 100

        If nDownPaym > lnDownPaym Then Return False

        Return True
    End Function

    Public Function Encode() As Boolean
        Dim lsUserxx As String = ""
        Dim lsSeries As String = ""
        Dim lsNameCd As String = ""
        Dim lsRand01 As String = ""
        Dim lsRand02 As String = ""
        Dim lsDownPy As String = ""

        Dim lnRand01 As Integer = 0
        Dim lnRand02 As Integer = 0

        sGOCASNox = ""

        If sTransNox = "" Then
            Message = "UNSET Transaction Number."
            Return False
        End If

        If sUserIDxx = "" Then
            Message = "UNSET App User ID."
            Return False
        End If

        If sClientNm = "" Then
            If sLastName = "" Then
                Message = "UNSET Last Name."
                Return False
            End If

            If sFrstName = "" Then
                Message = "UNSET First Name."
                Return False
            End If

            If sMiddName = "" Then
                Message = "UNSET Middle Name."
                Return False
            End If

            sClientNm = sLastName & ", " & sFrstName & " "
            If sSuffixNm <> "" Then sClientNm = sClientNm & sSuffixNm & " "
            sClientNm = sClientNm & sMiddName
        End If

        If cWithCIxx = "" Then
            Message = "UNSET value if CI was needed."
            Return False
        End If

        'process random number 1
        lnRand01 = Randomizer()
        lsRand01 = procRandom(lnRand01)

        'process random number 2
        lnRand02 = Randomizer()
        lsRand02 = procRandom(lnRand02)

        'process user id
        lsUserxx = procUser(lnRand01)

        'process series
        lsSeries = procSeries(lsUserxx, lnRand02)

        'process downpayment
        lsDownPy = procDownPayment(lnRand01, lnRand02)

        'process client
        lsNameCd = sClientNm
        If sClientNm.Length <> 3 Then lsNameCd = ClientCode(sLastName, sFrstName, sMiddName, sSuffixNm)
        lsNameCd = procName(lsNameCd, lnRand01, lnRand02)

        'user(XXXXX) + series(XXXXX) + name(XXXX) + rand1(X) + down(XX) + rand2(X)
        sGOCASNox = lsUserxx + lsSeries + lsNameCd + lsRand01 + lsDownPy + lsRand02

        Return True
    End Function

    Public Function Encode(ByVal fsGOCASNox As String, ByVal fnNewDownP As Integer, ByVal fbNeedCI As Boolean) As Boolean
        If (Not Decode(fsGOCASNox)) Then Return False

        'set the new downpayment
        nDownPaym = fnNewDownP
        'is CI needed?
        cWithCIxx = CStr(IIf(fbNeedCI = True, "1", "0"))

        If (Not Encode()) Then
            Message = "Unable to issue FINAL GOCAS Number."
            Return False
        End If

        Return True
    End Function

    Public Function Decode(ByVal fsGOCASNox As String) As Boolean
        Message = "Unable to decode the GOCAS Number."

        If fsGOCASNox.Length <> 15 And fsGOCASNox.Length <> 16 And fsGOCASNox.Length <> 17 And fsGOCASNox.Length <> 18 Then
            Message = Message + "Invalid code format detected."
            Return False
        End If

        fsGOCASNox = fsGOCASNox.ToUpper

        Dim lsUserxx As String = ""
        Dim lsSeries As String = ""
        Dim lsNameCd As String = ""
        Dim lsRand01 As String = ""
        Dim lsRand02 As String = ""
        Dim lsDownPy As String = ""

        Dim lnRand01 As Integer = 0
        Dim lnRand02 As Integer = 0
        Dim lnResult As Long = 0

        sGOCASNox = fsGOCASNox

        If fsGOCASNox.Length = 18 Then
            lsUserxx = sGOCASNox.Substring(0, 5) 'user id length is 10
            lsSeries = sGOCASNox.Substring(5, 5)
            lsNameCd = sGOCASNox.Substring(10, 4)
            lsRand01 = sGOCASNox.Substring(14, 1)
            lsDownPy = sGOCASNox.Substring(15, 2)
            lsRand02 = sGOCASNox.Substring(17, 1)
        ElseIf fsGOCASNox.Length = 17 Then
            lsUserxx = sGOCASNox.Substring(0, 4) 'user id length is 12
            lsSeries = sGOCASNox.Substring(4, 5)
            lsNameCd = sGOCASNox.Substring(9, 4)
            lsRand01 = sGOCASNox.Substring(13, 1)
            lsDownPy = sGOCASNox.Substring(14, 2)
            lsRand02 = sGOCASNox.Substring(16, 1)
        ElseIf fsGOCASNox.Length = 16 Then 'Name is Sunny, . .
            lsUserxx = sGOCASNox.Substring(0, 5)
            lsSeries = sGOCASNox.Substring(5, 5)
            lsNameCd = sGOCASNox.Substring(10, 2)
            lsRand01 = sGOCASNox.Substring(12, 1)
            lsDownPy = sGOCASNox.Substring(13, 2)
            lsRand02 = sGOCASNox.Substring(15, 1)
        ElseIf fsGOCASNox.Length = 15 Then 'Name is La Vera Mimosa Credit Corp, . .
            'user id is 10
            lsUserxx = sGOCASNox.Substring(0, 4)
            lsSeries = sGOCASNox.Substring(4, 5)
            lsNameCd = sGOCASNox.Substring(9, 2)
            lsRand01 = sGOCASNox.Substring(11, 1)
            lsDownPy = sGOCASNox.Substring(12, 2)
            lsRand02 = sGOCASNox.Substring(14, 1)
        End If

        'deserialize 1st random number
        lnRand01 = CInt(DeSerializeNumber(lsRand01))
        'deserialize 2nd random number
        lnRand02 = CInt(DeSerializeNumber(lsRand02))

        'deserialize user
        lnResult = DeSerializeNumber(lsUserxx) - lnRand01
        sUserIDxx = BRANCHCD + lnResult.ToString

        'deserialize transaction number
        sTransNox = SerializeNumber(lnResult)
        lnResult = DeSerializeNumber(lsSeries) - lnRand02
        sTransNox = sTransNox + lnResult.ToString

        'deseriaized downpayment
        If (Left(lsDownPy, 1) = "0") Then lsDownPy = lsDownPy.Substring(1)

        lnResult = DeSerializeNumber(lsDownPy)
        If lnResult > 1000 Then
            cWithCIxx = "1"
            lnResult = CLng(lnResult.ToString.Substring(1))
        Else
            cWithCIxx = "0"
        End If
        nDownPaym = CInt(lnResult - lnRand01 - lnRand02)

        'deserialize name
        lnResult = Radix2Dec(lsNameCd, 16)
        lnResult = lnResult - lnRand01 - lnRand02
        sClientNm = SerializeNumber(lnResult)

        Return True
    End Function

    Private Function procRandom(ByVal fnValue As Integer) As String
        Return SerializeNumber(fnValue) 'serialize value
    End Function

    Private Function procUser(ByVal fnRand01 As Integer) As String
        'get the long value of user id then add the first random number
        'serialize value
        Return SerializeNumber(Long.Parse(sUserIDxx.Substring(4)) + fnRand01) 'serialize value
    End Function

    Private Function procSeries(ByVal fsUser As String, ByVal fnRand02 As Integer) As String
        'get the long value of transaction number and add the second random number
        'serialize value
        Return SerializeNumber(Long.Parse(sTransNox.Substring(fsUser.Length())) + fnRand02)
    End Function

    Private Function procDownPayment(ByVal fnRand01 As Integer, ByVal fnRand02 As Integer) As String
        Dim lsDownPy As String

        'add string char of WithCI to the string value of downpayment
        lsDownPy = cWithCIxx & (nDownPaym + fnRand01 + fnRand02).ToString.PadLeft(3, CChar("0"))

        'serialize value
        lsDownPy = SerializeNumber(Long.Parse(lsDownPy))

        'if the result is only 1 character, pad it left with "0"
        If (lsDownPy.Length <> 2) Then lsDownPy = lsDownPy.PadLeft(2, CChar("0"))

        Return lsDownPy
    End Function

    Private Function procName(ByVal fsCode As String, ByVal fnRand01 As Integer, ByVal fnRand02 As Integer) As String
        'deserialize the name code and add the 1st and 2nd random numbers
        'then convert it to hexadecimal value
        Return (Dec2Radix(DeSerializeNumber(fsCode) + fnRand01 + fnRand02, 16)).ToString
    End Function

    Private Sub initValues()
        sTransNox = ""
        sUserIDxx = ""
        sClientNm = ""
        cWithCIxx = ""
        sGOCASNox = ""

        sLastName = ""
        sFrstName = ""
        sMiddName = ""
        sSuffixNm = ""

        'default value 200 means cash
        nDownPaym = 200
    End Sub

    Public Sub New()
        initValues()
    End Sub
End Class
