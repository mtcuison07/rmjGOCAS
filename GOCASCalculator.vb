'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€
' Guanzon Software Engineering Group
' Guanzon Group of Companies
' Perez Blvd., Dagupan City
'
'   GOCAS Credit Score Calculator
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
'   Mac 2019.12.23 03:40 PM
'       Started creating this object.
'   Mac 2020.03.14 11:05 AM
'       Added individual points saving.
' ==========================================================================================
'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€
Imports ggcAppDriver
Imports Newtonsoft.Json.Linq

Public Class GOCASCalculator
    Private Const pxeCONTACT_INFO As Double = 20
    Private Const pxeRESIDENCE_INFO As Double = 40
    Private Const pxeDISBURSEMENT_INFO As Double = 40

    Private psCatInfox As String
    Private poApp As GRider

    Private pnContactx As Integer 'contact info score
    Private pnResidnce As Integer 'residence info score
    Private pnDisposbl As Integer 'disposable income score

    'contact info individual score
    Private pnMobilePt As Integer 'mobile category points
    Private pnCvilStat As Integer 'civil status points
    Private pnFBPoints As Integer 'facebook category points
    'disposable income individual score
    Private pnSelfEmpx As Integer
    Private pnEmployed As Integer
    Private pnFinancer As Integer
    Private pnPensionr As Integer
    Private pnDpndntPt As Integer

    Dim pnMaxContactInf As Integer
    Dim pnMaxResidenceI As Integer
    Dim pnMaxDisposable As Integer

    WriteOnly Property setAppDriver() As GRider
        Set(value As GRider)
            poApp = value
        End Set
    End Property

    WriteOnly Property setJSON() As String
        Set(value As String)
            psCatInfox = value
        End Set
    End Property

    ReadOnly Property getContactInfoRate() As Double
        Get
            Return (pnContactx / pnMaxContactInf) * pxeCONTACT_INFO
        End Get
    End Property
    ReadOnly Property getResidenceInfoRate() As Double
        Get
            Return (pnResidnce / pnMaxResidenceI) * pxeRESIDENCE_INFO
        End Get
    End Property
    ReadOnly Property getDisposableIncomeRate() As Double
        Get
            Return (pnDisposbl / pnMaxDisposable) * pxeDISBURSEMENT_INFO
        End Get
    End Property

    ReadOnly Property getContactInfoPoints() As Integer
        Get
            Return pnContactx
        End Get
    End Property
    ReadOnly Property getResidenceInfoPoints() As Integer
        Get
            Return pnResidnce
        End Get
    End Property
    ReadOnly Property getDisposableIncomePoints() As Integer
        Get
            Return pnDisposbl
        End Get
    End Property

    ReadOnly Property getMobilePoints() As Integer
        Get
            Return pnMobilePt
        End Get
    End Property
    ReadOnly Property getCivilStatPoints() As Integer
        Get
            Return pnCvilStat
        End Get
    End Property
    ReadOnly Property getFBPoints As Integer
        Get
            Return pnFBPoints
        End Get
    End Property

    ReadOnly Property getSelfEmployedPoints() As Integer
        Get
            Return pnSelfEmpx
        End Get
    End Property
    ReadOnly Property getEmployedPoints() As Integer
        Get
            Return pnEmployed
        End Get
    End Property
    ReadOnly Property getFinancedPoints() As Integer
        Get
            Return pnFinancer
        End Get
    End Property
    ReadOnly Property getPensionerPoints() As Integer
        Get
            Return pnPensionr
        End Get
    End Property
    ReadOnly Property getDependentsPoints() As Integer
        Get
            Return pnDpndntPt
        End Get
    End Property

    Public Function Compute() As Double
        computeContactInfo()
        computeDisposable()
        computeResidence()

        Debug.Print("CONTACT INFO SCORE: " + CStr(pnContactx))
        Debug.Print("RESIDENCE INFO SCORE: " + CStr(pnResidnce))
        Debug.Print("DISPOSABLE INCOME SCORE: " + CStr(pnDisposbl))

        pnMaxContactInf = getConfigValue("gocas.ms.contact_info")
        pnMaxResidenceI = getConfigValue("gocas.ms.residence_info")
        pnMaxDisposable = getConfigValue("gocas.ms.disbursement_info")

        Dim lnContactx As Double = (pnContactx / pnMaxContactInf) * pxeCONTACT_INFO
        Dim lnResidnce As Double = (pnResidnce / pnMaxResidenceI) * pxeRESIDENCE_INFO
        Dim lnDisposbl As Double = (pnDisposbl / pnMaxDisposable) * pxeDISBURSEMENT_INFO

        'mac 2020.06.18
        lnContactx = CDbl(IIf(lnContactx >= pxeCONTACT_INFO, pxeCONTACT_INFO, lnContactx))
        lnResidnce = CDbl(IIf(lnResidnce >= pxeRESIDENCE_INFO, pxeRESIDENCE_INFO, lnResidnce))
        lnDisposbl = CDbl(IIf(lnDisposbl >= pxeDISBURSEMENT_INFO, pxeDISBURSEMENT_INFO, lnDisposbl))

        If lnContactx + lnResidnce + lnDisposbl < 0 Then
            Return 0
        Else
            Return lnContactx + lnResidnce + lnDisposbl
        End If
    End Function

    Private Function computeDisposable() As Boolean
        Dim lsValue As String
        Dim loJSON As JObject
        Dim loJSON1 As JObject
        Dim loArray As JArray

        pnDisposbl = 0

        'get MEANS INFO KEY
        loJSON = CType(JObject.Parse(psCatInfox)("means_info"), JObject)

        'PROCESS EMPLOYED MEANS
        pnEmployed = 0
        loJSON1 = CType(JObject.Parse(loJSON.ToString)("employed"), JObject)
        'employment sector must not be empty
        If CStr(loJSON1.GetValue("cEmpSectr")) <> "" Then
            Select Case CStr(loJSON1.GetValue("cEmpSectr"))
                Case "0"
                    'score when applicant is public employee
                    lsValue = "gocas.cs.means_info.cIncmeSrc.employed.cEmpSectr.eq.0"
                    pnEmployed = pnEmployed + getConfigValue(lsValue)
                    Debug.Print(lsValue)

                    'score when public employee applicant is a uniformed personnel
                    lsValue = "gocas.cs.means_info.cIncmeSrc.employed.cEmpSectr.eq.0.cUniforme.eq." & CStr(loJSON1.GetValue("cUniforme"))
                    pnEmployed = pnEmployed + getConfigValue(lsValue)
                    Debug.Print(lsValue)

                    'Score when public employee applicant is a military personnel(AFP)
                    lsValue = "gocas.cs.means_info.cIncmeSrc.employed.cEmpSectr.eq.0.cMilitary.eq." & CStr(loJSON1.GetValue("cMilitary"))
                    pnEmployed = pnEmployed + getConfigValue(lsValue)
                    Debug.Print(lsValue)

                    'score for position
                    lsValue = "gocas.cs.means_info.cIncmeSrc.employed.cEmpSectr.eq.0.cEmpLevlx.eq." & CStr(loJSON1.GetValue("cEmpLevlx"))
                    pnEmployed = pnEmployed + getConfigValue(lsValue)
                    Debug.Print(lsValue)

                    'score for tenure/years
                    lsValue = ""
                    If CDbl(CStr(loJSON1.GetValue("nLenServc"))) < 1 Then
                        'mac 2020.06.18
                        ' comment this code since a category of below/above 6 months is added
                        'lsValue = "gocas.cs.means_info.cIncmeSrc.employed.cEmpSectr.eq." & CStr(loJSON1.GetValue("cEmpSectr")) & ".nLenServc.lt.1"

                        'mac 2020.06.18
                        If CDbl(CStr(loJSON1.GetValue("nLenServc"))) < 0.5 Then'less than 6 mont hs
                            lsValue = "gocas.cs.means_info.cIncmeSrc.employed.cEmpSectr.eq.0.nLenServc.lt.1/2"
                        Else 'greater than 6 months
                            lsValue = "gocas.cs.means_info.cIncmeSrc.employed.cEmpSectr.eq.0.nLenServc.mteq.1/2"
                        End If
                    ElseIf CDbl(CStr(loJSON1.GetValue("nLenServc"))) < 3 Then
                        lsValue = "gocas.cs.means_info.cIncmeSrc.employed.cEmpSectr.eq." & CStr(loJSON1.GetValue("cEmpSectr")) & ".nLenServc.lt.3"
                    ElseIf CDbl(CStr(loJSON1.GetValue("nLenServc"))) >= 3 Then
                        lsValue = "gocas.cs.means_info.cIncmeSrc.employed.cEmpSectr.eq." & CStr(loJSON1.GetValue("cEmpSectr")) & ".nLenServc.mteq.3"
                    End If
                    pnEmployed = pnEmployed + getConfigValue(lsValue)
                    Debug.Print(lsValue)
                Case "1"
                    'score for company level
                    lsValue = "gocas.cs.means_info.cIncmeSrc.employed.cEmpSectr.eq.1.cCompLevl.eq." & CStr(loJSON1.GetValue("cCompLevl"))
                    pnEmployed = pnEmployed + getConfigValue(lsValue)
                    Debug.Print(lsValue)

                    'nature of business
                    If CStr(loJSON1.GetValue("sIndstWrk")) <> "" Then
                        lsValue = "gocas.cs.means_info.cIncmeSrc.employed.cEmpSectr.eq.1.sIndstWrk"
                        pnEmployed = pnEmployed + getConfigValue(lsValue)
                        Debug.Print(lsValue)
                    End If

                    'score for position
                    lsValue = "gocas.cs.means_info.cIncmeSrc.employed.cEmpSectr.eq.1.cEmpLevlx.eq." & CStr(loJSON1.GetValue("cEmpLevlx"))
                    pnEmployed = pnEmployed + getConfigValue(lsValue)
                    Debug.Print(lsValue)

                    'score for tenure/years
                    lsValue = ""
                    If CDbl(CStr(loJSON1.GetValue("nLenServc"))) < 1 Then
                        lsValue = "gocas.cs.means_info.cIncmeSrc.employed.cEmpSectr.eq." & CStr(loJSON1.GetValue("cEmpSectr")) & ".nLenServc.lt.1"
                    ElseIf CDbl(CStr(loJSON1.GetValue("nLenServc"))) < 3 Then
                        lsValue = "gocas.cs.means_info.cIncmeSrc.employed.cEmpSectr.eq." & CStr(loJSON1.GetValue("cEmpSectr")) & ".nLenServc.lt.3"
                    ElseIf CDbl(CStr(loJSON1.GetValue("nLenServc"))) >= 3 Then
                        lsValue = "gocas.cs.means_info.cIncmeSrc.employed.cEmpSectr.eq." & CStr(loJSON1.GetValue("cEmpSectr")) & ".nLenServc.mteq.3"
                    End If
                    pnEmployed = pnEmployed + getConfigValue(lsValue)
                    Debug.Print(lsValue)
                Case "2"
                    If CStr(loJSON1.GetValue("cOFWRegnx")) <> "" Then
                        If CStr(loJSON1.GetValue("cOFWRegnx")) = "3" Then 'asia
                            lsValue = "gocas.cs.means_info.cIncmeSrc.employed.cEmpSectr.eq.2.cOFWRegnx.eq.3"
                            pnEmployed = pnEmployed + getConfigValue(lsValue)
                            Debug.Print(lsValue)

                            'work category
                            lsValue = "gocas.cs.means_info.cIncmeSrc.employed.cEmpSectr.eq.2.cOcCatgry.eq." & CStr(loJSON1.GetValue("cOcCatgry"))
                            pnEmployed = pnEmployed + getConfigValue(lsValue)
                            Debug.Print(lsValue)
                        Else 'other region
                            lsValue = "gocas.cs.means_info.cIncmeSrc.employed.cEmpSectr.eq.2.cOFWRegnx.neq.3"
                            pnEmployed = pnEmployed + getConfigValue(lsValue)
                            Debug.Print(lsValue)
                        End If
                    End If

                    'score for tenure/years
                    lsValue = ""
                    If CDbl(CStr(loJSON1.GetValue("nLenServc"))) < 1 Then
                        lsValue = "gocas.cs.means_info.cIncmeSrc.employed.cEmpSectr.eq." & CStr(loJSON1.GetValue("cEmpSectr")) & ".nLenServc.lt.1"
                    ElseIf CDbl(CStr(loJSON1.GetValue("nLenServc"))) < 3 Then
                        lsValue = "gocas.cs.means_info.cIncmeSrc.employed.cEmpSectr.eq." & CStr(loJSON1.GetValue("cEmpSectr")) & ".nLenServc.lt.3"
                    ElseIf CDbl(CStr(loJSON1.GetValue("nLenServc"))) >= 3 Then
                        lsValue = "gocas.cs.means_info.cIncmeSrc.employed.cEmpSectr.eq." & CStr(loJSON1.GetValue("cEmpSectr")) & ".nLenServc.mteq.3"
                    End If
                    pnEmployed = pnEmployed + getConfigValue(lsValue)
                    Debug.Print(lsValue)
            End Select
        End If
        '***END PROCESS EMPLOYED MEANS

        'PROCESS PENSIONER MEANS
        pnPensionr = 0
        loJSON1 = CType(JObject.Parse(loJSON.ToString)("pensioner"), JObject)
        lsValue = "gocas.cs.means_info.cIncmeSrc.pensioner.cPenTypex.eq." & CStr(loJSON1.GetValue("cPenTypex"))
        pnPensionr = pnPensionr + getConfigValue(lsValue)
        Debug.Print(lsValue)
        '***END PROCESS PENSIONER MEANS

        'PROCESS FINANCED MEANS
        pnFinancer = 0
        loJSON1 = CType(JObject.Parse(loJSON.ToString)("financed"), JObject)
        lsValue = "gocas.cs.means_info.cIncmeSrc.financed.eq." & CStr(loJSON1.GetValue("sReltnCde"))
        pnFinancer = pnFinancer + getConfigValue(lsValue)
        Debug.Print(lsValue)
        '***END PROCESS FINANCED MEANS

        'PROCESS SELF EMPLOYED MEANS
        pnSelfEmpx = 0
        loJSON1 = CType(JObject.Parse(loJSON.ToString)("self_employed"), JObject)

        'ownership size/capital
        lsValue = "gocas.cs.means_info.cIncmeSrc.self_employed.cOwnSizex.eq." & CStr(loJSON1.GetValue("cOwnSizex"))
        pnSelfEmpx = pnSelfEmpx + getConfigValue(lsValue)
        Debug.Print(lsValue)

        'length of business
        lsValue = ""
        If CInt(CStr(loJSON1.GetValue("nBusLenxx"))) < 1 Then
            lsValue = "gocas.cs.means_info.cIncmeSrc.self_employed.nBusLenxx.lt.1"
        ElseIf CInt(CStr(loJSON1.GetValue("nBusLenxx"))) < 3 Then
            lsValue = "gocas.cs.means_info.cIncmeSrc.self_employed.nBusLenxx.lt.3"
        ElseIf CInt(CStr(loJSON1.GetValue("nBusLenxx"))) >= 3 Then
            lsValue = "gocas.cs.means_info.cIncmeSrc.self_employed.nBusLenxx.mteq.3"
        End If
        pnSelfEmpx = pnSelfEmpx + getConfigValue(lsValue)
        Debug.Print(lsValue)

        'type of company
        lsValue = "gocas.cs.means_info.cIncmeSrc.self_employed.cOwnTypex.eq." & CStr(loJSON1.GetValue("cBusTypex"))
        pnSelfEmpx = pnSelfEmpx + getConfigValue(lsValue)
        Debug.Print(lsValue)

        'nature of business
        If CStr(loJSON1.GetValue("sIndstBus")) <> "" Then
            lsValue = "gocas.cs.means_info.cIncmeSrc.self_employed.sIndstBus"
            pnSelfEmpx = pnSelfEmpx + getConfigValue(lsValue)
            Debug.Print(lsValue)
        End If
        '***END PROCESS SELF EMPLOYED MEANS

        'get DISBURSEMENT INFO KEY
        loJSON = CType(JObject.Parse(psCatInfox)("disbursement_info"), JObject)

        'PROCESS DEPENDENT INFO
        loJSON1 = CType(JObject.Parse(loJSON.ToString)("dependent_info"), JObject)
        loArray = JArray.Parse(loJSON1.GetValue("children").ToString)

        For lnCtr = 0 To loArray.Count - 1
            If CStr(loArray(lnCtr)("sRelatnCD")) = "0" Then
                'is the dependent has work? where is he working?
                lsValue = "gocas.cs.disbursement_info.dependent_info.sRelatnCD.eq.0.cHasWorkx.eq." & CStr(loArray(lnCtr)("cHasWorkx")) & _
                            ".cWorkType.eq." & CStr(loArray(lnCtr)("cWorkType"))
                pnSelfEmpx = pnSelfEmpx + getConfigValue(lsValue)
                Debug.Print(lsValue)

                'is the dependent studying? at what level?
                lsValue = "gocas.cs.disbursement_info.dependent_info.sRelatnCD.eq.0.cIsPupilx.eq." & CStr(loArray(lnCtr)("cIsPupilx")) & _
                            ".sEducLevl.eq." & CStr(loArray(lnCtr)("sEducLevl"))
                pnSelfEmpx = pnSelfEmpx + getConfigValue(lsValue)
                Debug.Print(lsValue)

                'is the dependent studying? is he on private school?
                lsValue = "gocas.cs.disbursement_info.dependent_info.sRelatnCD.eq.0.cIsPupilx.eq." & CStr(loArray(lnCtr)("cIsPupilx")) & _
                            ".cIsPrivte.eq." & CStr(loArray(lnCtr)("cIsPrivte"))
                pnSelfEmpx = pnSelfEmpx + getConfigValue(lsValue)
                Debug.Print(lsValue)

                'is the dependent studying? is he scholar?
                lsValue = "gocas.cs.disbursement_info.dependent_info.sRelatnCD.eq.0.cIsPupilx.eq." & CStr(loArray(lnCtr)("cIsPupilx")) & _
                            ".cIsSchlrx.eq." & CStr(loArray(lnCtr)("cIsSchlrx"))
                pnSelfEmpx = pnSelfEmpx + getConfigValue(lsValue)
                Debug.Print(lsValue)
            Else
                lsValue = "gocas.cs.disbursement_info.dependent_info.sRelatnCD.eq." & CStr(loArray(lnCtr)("sRelatnCD")) & _
                            ".cDependnt.eq." & CStr(loArray(lnCtr)("cDependnt"))
                pnSelfEmpx = pnSelfEmpx + getConfigValue(lsValue)
                Debug.Print(lsValue)
            End If
        Next

        pnDisposbl = pnSelfEmpx + pnEmployed + pnFinancer + pnPensionr
        Debug.Print("DISPOSABLE INCOME SCORE: " + CStr(pnDisposbl))
        Return True
    End Function

    Private Function computeResidence() As Boolean
        Dim lsValue As String
        Dim loJSON As JObject

        pnResidnce = 0

        'get RESIDENCE INFO KEY
        loJSON = CType(JObject.Parse(psCatInfox)("residence_info"), JObject)

        Select Case CStr(loJSON.GetValue("cOwnershp"))
            Case "0" 'owned
                'score on house ownership household
                lsValue = "gocas.cs.residence_info.cOwnershp.eq.0.cOwnOther.eq." & _
                            CStr(loJSON.GetValue("cOwnOther"))
                pnResidnce = pnResidnce + getConfigValue(lsValue)
                Debug.Print(lsValue)

                'score for house type
                lsValue = "gocas.cs.residence_info.cOwnershp.eq.0.cOwnOther.eq." & _
                            CStr(loJSON.GetValue("cOwnOther")) & _
                            ".cHouseTyp.eq." & _
                            CStr(loJSON.GetValue("cHouseTyp"))
                pnResidnce = pnResidnce + getConfigValue(lsValue)
                Debug.Print(lsValue)
            Case "1", "2" 'rented, caretaker
                'score on house ownership
                lsValue = "gocas.cs.residence_info.cOwnershp.eq." & CStr(loJSON.GetValue("cOwnershp"))
                pnResidnce = pnResidnce + getConfigValue(lsValue)
                Debug.Print(lsValue)

                'score for house type
                lsValue = "gocas.cs.residence_info.cOwnershp.eq." & CStr(loJSON.GetValue("cOwnershp")) & _
                            ".cHouseTyp.eq." & CStr(loJSON.GetValue("cHouseTyp"))
                pnResidnce = pnResidnce + getConfigValue(lsValue)
                Debug.Print(lsValue)

                'score for length of stay
                lsValue = ""
                If CInt(loJSON.GetValue("cHouseTyp")) < 2 Then
                    lsValue = "gocas.cs.residence_info.cOwnershp.eq." & CStr(loJSON.GetValue("cOwnershp")) & _
                                ".nLenStayx.lt.2"
                Else
                    lsValue = "gocas.cs.residence_info.cOwnershp.eq." & CStr(loJSON.GetValue("cOwnershp")) & _
                                ".nLenStayx.mteq.2"
                End If
                pnResidnce = pnResidnce + getConfigValue(lsValue)
                Debug.Print(lsValue)
        End Select

        'is the house has garage
        lsValue = "gocas.cs.residence_info.cGaragexx.eq." & CStr(loJSON.GetValue("cGaragexx"))
        pnResidnce = pnResidnce + getConfigValue(lsValue)
        Debug.Print(lsValue)

        Debug.Print("RESIDENCE INFO SCORE: " + CStr(pnResidnce))
        Return True
    End Function

    Private Function computeContactInfo() As Boolean
        Dim lsValue As String
        Dim loJSON As JObject
        Dim loJSON1 As JObject
        Dim loArray As JArray
        Dim lnCtr As Integer

        pnContactx = 0

        'get APPLICANT INFO KEY
        loJSON = CType(JObject.Parse(psCatInfox)("applicant_info"), JObject)

        'PROCESS MOBILE NUMBER
        pnMobilePt = 0
        loArray = JArray.Parse(loJSON.GetValue("mobile_number").ToString)

        If loArray.Count > 0 Then
            Dim postpaidctr As Integer = 0
            Dim mptypescore As Integer = 0
            Dim mpagescore As Integer = 0
            Dim mppostpaid As Integer

            For lnCtr = 0 To loArray.Count - 1
                If CStr(loArray.Item(lnCtr)("sMobileNo")) = "" Then GoTo nextMobile

                'is the mobile no postpaid?
                If CStr(loArray.Item(lnCtr)("cPostPaid")) = "1" Then
                    'set type score
                    lsValue = "gocas.cs.applicant_info.mobile_number.cpostpaid.eq.1"
                    mptypescore = getConfigValue(lsValue)
                    Debug.Print(lsValue)

                    'is postpaid year is > 1?
                    If CInt(loArray.Item(lnCtr)("nPostYear")) > 1 And mpagescore = 0 Then
                        lsValue = "gocas.cs.applicant_info.mobile_number.cpostpaid.eq.1.npostyear.mt.1"
                        mpagescore = getConfigValue(lsValue)
                        Debug.Print(lsValue)
                    End If

                    postpaidctr = postpaidctr + 1

                    'is postpaid count is > 1?
                    If postpaidctr > 1 And mppostpaid = 0 Then
                        lsValue = "gocas.cs.applicant_info.mobile_number.cpostpaid.eq.1.count.mt.1"
                        mppostpaid = getConfigValue(lsValue)
                        Debug.Print(lsValue)
                    End If
                Else 'prepaid
                    If mptypescore = 0 Then
                        'set type score
                        lsValue = "gocas.cs.applicant_info.mobile_number.cpostpaid.eq.0"
                        mptypescore = getConfigValue(lsValue)
                        Debug.Print(lsValue)
                    End If
                End If
nextMobile:
            Next
            'total + type + quantity + age
            pnMobilePt = mptypescore + mppostpaid + mpagescore 'total mobile points

            pnContactx = pnContactx + pnMobilePt 'accumulate points
        End If
        '***END - PROCESS MOBILE NUMBER

        'PROCESS CIVIL STATUS
        pnCvilStat = 0
        lsValue = "gocas.cs.applicant_info.ccvilstat.eq." + CStr(loJSON.GetValue("cCvilStat"))
        pnCvilStat = getConfigValue(lsValue) 'civil status points
        pnContactx = pnContactx + pnCvilStat 'accumulate points
        Debug.Print(lsValue)
        '***END - PROCESS CIVIL STATUS

        'PROCESS FACEBOOK
        pnFBPoints = 0
        loJSON1 = JObject.Parse(loJSON.GetValue("facebook").ToString)
        If Trim(CStr(loJSON1.GetValue("sFBAcctxx"))) <> "" Then
            'add score having an FB account
            lsValue = "gocas.cs.applicant_info.facebook.sfbacctxx.neq.empty"
            pnFBPoints = pnFBPoints + getConfigValue(lsValue) 'add to fb points
            Debug.Print(lsValue)

            'add score if account is active
            lsValue = "gocas.cs.applicant_info.facebook.cacctstat.eq." & CStr(loJSON1.GetValue("cAcctStat"))
            pnFBPoints = pnFBPoints + getConfigValue(lsValue)
            Debug.Print(lsValue)

            If CInt(CStr(loJSON1.GetValue("nNoFriend"))) > 100 Then
                'add score if friends is more than 100
                lsValue = "gocas.cs.applicant_info.facebook.nnofriend.mt.100"
                pnFBPoints = pnFBPoints + getConfigValue(lsValue) 'add to fb points
                Debug.Print(lsValue)
            End If

            If CInt(CStr(loJSON1.GetValue("nYearxxxx"))) > 1 Then
                'add score if account age is more than 1 year
                lsValue = "gocas.cs.applicant_info.facebook.nyearxxxx.mt.1"
                pnFBPoints = pnFBPoints + CInt(getConfigValue(lsValue)) 'add to fb points
                Debug.Print(lsValue)
            End If

            pnContactx = pnContactx + pnFBPoints 'accumulate points
        End If

        Debug.Print("CONTACT INFO SCORE: " + CStr(pnContactx))
        Return True
    End Function

    Public Sub New()
        Call initRecord()

        poApp = Nothing
    End Sub

    Private Sub initRecord()
        psCatInfox = ""

        pnContactx = 0
        pnResidnce = 0
        pnDisposbl = 0

        'contact info individual score
        pnMobilePt = 0 'mobile category points
        pnCvilStat = 0 'civil status points
        pnFBPoints = 0 'facebook category points
        'disposable income individual score
        pnSelfEmpx = 0 'self employed points
        pnEmployed = 0 'employed points
        pnFinancer = 0 'with financer points
        pnPensionr = 0 'pensioner points
        pnDpndntPt = 0 'dependent points
    End Sub

    Private Function getConfigValue(ByVal fsConfigID As String) As Integer
        If TypeName(poApp) = "Nothing" Then Return 0

        Dim lsSQL As String = "SELECT IFNULL(b.sConfigVl, a.sConfigVl) sConfigVl" & _
                                " FROM xxxSysConfig a" & _
                                    " LEFT JOIN xxxSysConfigHistory b" & _
                                        " ON a.sConfigCd = b.sConfigCd" & _
                                            " AND " + dateParm(poApp.SysDate) + " BETWEEN b.dDateFrom AND b.dDateThru" & _
                                " WHERE a.sConfigCD = " & strParm(fsConfigID) & _
                                    " AND a.dDteAdded <= " & dateParm(poApp.SysDate)

        Dim loDT As DataTable = poApp.ExecuteQuery(lsSQL)

        If loDT.Rows.Count = 0 Then Return 0

        Return CInt(loDT(0)("sConfigVl"))
    End Function
End Class