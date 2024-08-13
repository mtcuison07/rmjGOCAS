Imports Newtonsoft
Imports Newtonsoft.Json.Linq
Imports ggcAppDriver

Module modMain
    Private p_oApp As New GRider("GRider")

    Private Sub testCalc()
        Dim instance As New GOCASCalculator
        Dim lsValue As String

        lsValue = "{'sBranchCd':'M111','dAppliedx':'2020-09-02','sClientNm':'legaspi, wilter john tan','cUnitAppl':'0','sUnitAppl':'HONDA','nDownPaym':'20000','dCreatedx':'2020-09-02 11:22:44','cApplType':'1','sModelIDx':'M00120002','nAcctTerm':'36','nMonAmort':'6118','dTargetDt':'','applicant_info':{'sLastName':'legaspi','sFrstName':'wilter john','sSuffixNm':'','sMiddName':'tan','sNickName':'','dBirthDte':'1994-12-12','sBirthPlc':'0314','sCitizenx':'01','mobile_number':[{'sMobileNo':'09156902719','cPostPaid':'0','nPostYear':'0'}],'landline':[{'sPhoneNox':''}],'cCvilStat':'1','cGenderCd':'0','sMaidenNm':'palaganas','email_address':[{'sEmailAdd':'wilterjohn_legaspi@yahoo.com'}],'facebook':{'sFBAcctxx':'wilter john t. legaspi','cAcctStat':'1','nNoFriend':'0','nYearxxxx':'0'},'sVibeAcct':''},'residence_info':{'cOwnershp':'0','cOwnOther':'2','rent_others':{},'sCtkReltn':'','cHouseTyp':'0','cGaragexx':'1','present_address':{'sLandMark':'Vape Shop Sa Kanto Ng Nava','sHouseNox':'253','sAddress1':'','sAddress2':'Nava','sTownIDxx':'0314','sBrgyIDxx':'1100139'},'permanent_address':{'sLandMark':'Vape Shop Sa Kanto Ng Nava','sHouseNox':'253','sAddress1':'','sAddress2':'Nava','sTownIDxx':'0314','sBrgyIDxx':'1100139'}},'means_info':{'cIncmeSrc':'1','employed':{'cEmpSectr':'','sIndstWrk':'','sEmployer':'','sWrkAddrx':'','sWrkTownx':'','sPosition':'','sFunction':'','cEmpStatx':'','nLenServc':'0','nSalaryxx':'0','sWrkTelno':''},'self_employed':{'sIndstBus':'Business Services','sBusiness':'Vape Shop','sBusAddrx':'Tapuac Dist. ','sBusTownx':'0314','cBusTypex':'0','nBusLenxx':'1','nBusIncom':'50000','nMonExpns':'8000','cOwnSizex':'1'},'pensioner':{'cPenTypex':'','nPensionx':'0','nRetrYear':'0'},'financed':{'sReltnCde':'','sFinancer':'','nEstIncme':'0','sNatnCode':'','sMobileNo':'','sFBAcctxx':'','sEmailAdd':''},'other_income':{'nOthrIncm':'','sOthrIncm':''}},'other_info':{'sUnitUser':'0','sPurposex':'0','sUnitPayr':'0','sSrceInfo':'Facebook','personal_reference':[{'sRefrNmex':'mam ruby lo','sRefrMPNx':'09365105011','sRefrAddx':'perez','sRefrTown':'0314'},{'sRefrNmex':'david kale palaganas','sRefrMPNx':'09358888986','sRefrAddx':'tapuac dist.','sRefrTown':'0314'},{'sRefrNmex':'teresita palaganas','sRefrMPNx':'09365105011','sRefrAddx':'tapuac dist.','sRefrTown':'0314'}]},'comaker_info':{'sLastName':'','sFrstName':'','sSuffixNm':'','sMiddName':'','sNickName':'','dBirthDte':'','sBirthPlc':'','cIncmeSrc':'','sReltnCde':'','mobile_number':[],'sFBAcctxx':''},'spouse_info':{'personal_info':{'sLastName':'legaspi','sFrstName':'camille','sSuffixNm':'','sMiddName':'arenas','sNickName':'','dBirthDte':'1995-06-06','sBirthPlc':'0314','sCitizenx':'01','mobile_number':[{'sMobileNo':'09959210057','cPostPaid':'0','nPostYear':'0'}],'landline':[{'sPhoneNox':''}],'cCvilStat':'1','cGenderCd':'1','sMaidenNm':'','email_address':[{'sEmailAdd':''}],'facebook':{'sFBAcctxx':'camille','cAcctStat':'','nNoFriend':'0','nYearxxxx':'0'},'sVibeAcct':''},'residence_info':{'cOwnershp':'','rent_others':{},'sCtkReltn':'','cHouseTyp':'','cGaragexx':'','present_address':{'sLandMark':'vape shop sa mismo kanto nava','sHouseNox':'253','sAddress1':'','sAddress2':'nava','sTownIDxx':'0314','sBrgyIDxx':'1100139'},'permanent_address':{'sLandMark':'vape shop sa mismo kanto nava','sHouseNox':'253','sAddress1':'','sAddress2':'nava','sTownIDxx':'0314','sBrgyIDxx':'1100139'}}},'spouse_means':{'cIncmeSrc':'1','employed':{'cEmpSectr':'','sIndstWrk':'','sEmployer':'','sWrkAddrx':'','sWrkTownx':'','sPosition':'','sFunction':'','cEmpStatx':'','nLenServc':'0','nSalaryxx':'0','sWrkTelno':''},'self_employed':{'sIndstBus':'Business Services','sBusiness':'Vape Shop','sBusAddrx':'Nava Tapuac Dist.','sBusTownx':'0314','cBusTypex':'0','nBusLenxx':'1','nBusIncom':'50000','nMonExpns':'8000','cOwnSizex':'0'},'pensioner':{'cPenTypex':'','nPensionx':'0','nRetrYear':'0'},'financed':{'sReltnCde':'','sFinancer':'','nEstIncme':'0','sNatnCode':'','sMobileNo':'','sFBAcctxx':'','sEmailAdd':''},'other_income':{'nOthrIncm':'','sOthrIncm':''}},'disbursement_info':{'dependent_info':{'nHouseHld':'0','children':[{}]},'properties':{'sProprty1':'','sProprty2':'','sProprty3':'','cWith4Whl':'','cWith3Whl':'','cWith2Whl':'','cWithRefx':'','cWithTVxx':'','cWithACxx':''},'monthly_expenses':{'nElctrcBl':'7000','nWaterBil':'500','nFoodAllw':'5000','nLoanAmtx':'0'},'bank_account':{'sBankName':'','sAcctType':'0'},'credit_card':{'sBankName':'','nCrdLimit':'0','nSinceYrx':'0'}}}"

        instance.setAppDriver = p_oApp
        instance.setJSON = lsValue
        MsgBox(instance.Compute())
    End Sub

    Sub Main()
        'testCalc()

        Dim instance As New GOCASCodeGen

        If (instance.Decode("CI4IT1FFVI8ABEC67B")) Then
            Debug.Print("TRAN NO: " + instance.TransactionNo)
            Debug.Print("USER ID: " + instance.UserID)
            Debug.Print("WITH CI: " + Str(instance.IsCINeeded))
            Debug.Print("APRV DP: " + Str(instance.DownPayment))
        Else
            MsgBox(instance.Message)
        End If

        ''instance.Encode("CI5AF190DW92FEFY5E", 48, True)
        ' ''Debug.Print(instance.GOCASApprvl)
        'MsgBox("tapos na")



        'p_oApp = New GRider("GRider")

        'If Not p_oApp.LoadEnv() Then
        '    MsgBox("Unable to load GhostRider!")
        '    Exit Sub
        'End If
        'If Not p_oApp.LogUser("M001111122") Then
        '    MsgBox("User unable to log!")
        '    Exit Sub
        'End If

        'While (0 <> 1)
        '    Dim instance As New GOCASCodeGen

        '    'instance.Decode("BB8I114Q288A4282RB")
        '    ''------------------------------------------------------------------
        instance.UserID = "GAP021001049" 'created
        instance.TransactionNo = "CI4IH2400019" 'table transaction number
        instance.LastName = "Martinez"
        instance.FirstName = "Dennis"
        instance.MiddleName = "Sabalburo"
        instance.SuffixName = ""
        instance.IsCINeeded = False 'is CI needed
        instance.DownPayment = 200 'approved downpayment
        instance.Encode() 'generate code
        Debug.Print(instance.GOCASApprvl)
        MsgBox(instance.GOCASApprvl)

        '    Dim lsSQL As String = instance.GOCASApprvl 'get code
        '    Debug.Print("FIRST GOCAS: " + lsSQL)
        '    MsgBox(lsSQL, , "FIRST GOCAS")

        '    ''here is how to validate GOCASNo
        '    ''   (GOCASNo, ClientNm, PNValuex, DownPaym)
        '    'MsgBox(instance.IsValidGOCASNo(lsSQL, _
        '    '                               "Cuison, Michael Jr. Bautista", _
        '    '                               100000, _
        '    '                               50000))
        '    ''------------------------------------------------------------------

        '    ''if there is an appeal, here is how to generate new GOCASNo based on old GOCASNo and new DP
        '    'instance.Encode(lsSQL, 40, False) 'generate code (OLD_GOCAS, NEW_DP)
        '    'lsSQL = instance.GOCASApprvl 'get code
        '    'Debug.Print("FINAL GOCAS: " + lsSQL)
        '    'MsgBox(lsSQL, , "FINAL GOCAS")

        '    'here is how to validate GOCASNo
        '    '   (GOCASNo, ClientNm, PNValuex, DownPaym)
        '    MsgBox(instance.IsValidGOCASNo("BB8HT1900005", _
        '                                   "BB8I514Q2B8A49C1UE", _
        '                                   "Cuison", _
        '                                   "Michael", _
        '                                   "Bautista", _
        '                                   "Jr.", _
        '                                   100000, _
        '                                   39999))
        '    '------------------------------------------------------------------
        'End While
    End Sub

    Public Function SerializeNumber(ByVal fnValue As Long) As String
        Return Dec2Radix(fnValue, 36)
    End Function

    Public Function DeSerializeNumber(ByVal fsValue As String) As Long
        Return Radix2Dec(fsValue, 36)
    End Function

    Public Function Randomizer() As Integer
        Return RandInRange(8, 16)
    End Function

    Public Function ClientCode(ByVal fsValue As String) As String
        Dim lasSplit() As String
        Dim lsValue As String
        Dim lnCtr As Integer
        Dim lsLast As String = ""
        Dim lsFirst As String = ""
        Dim lsMiddle As String = ""

        lsValue = ""
        lasSplit = CType(GetSplitedName(fsValue), String())


        For lnCtr = 0 To UBound(lasSplit)
            lasSplit(lnCtr) = Replace(lasSplit(lnCtr), ".", "") 'replace period in name to empty string
            lasSplit(lnCtr) = Replace(lasSplit(lnCtr), ",", "") 'replace comma in name to empty string

            'lsValue = lsValue + Right(lasSplit(lnCtr), 1).ToUpper

            Select Case lnCtr
                Case 0
                    lsLast = Right(lasSplit(lnCtr), 1).ToUpper
                Case 1
                    lsFirst = Right(lasSplit(lnCtr), 1).ToUpper
                Case 2
                    lsMiddle = Right(lasSplit(lnCtr), 1).ToUpper
            End Select
        Next

        lsValue = lsFirst + lsMiddle + lsLast

        Return lsValue
    End Function

    Public Function ClientCode(ByVal sLastName As String, _
                               ByVal sFrstName As String, _
                               ByVal sMiddName As String, _
                               ByVal sSuffixNm As String) As String

        Dim lsValue As String = ""
        ClientCode = lsValue

        sLastName = Trim(sLastName)
        sFrstName = Trim(sFrstName)
        sMiddName = sMiddName 'Trim(sMiddName)
        sSuffixNm = Trim(sSuffixNm)

        If sLastName = "" Then GoTo endProc
        If sFrstName = "" Then GoTo endProc
        If sMiddName = "" Then GoTo endProc

        sSuffixNm = sSuffixNm.Replace(".", "").Replace(",", "")

        If sSuffixNm <> "" Then
            lsValue = lsValue & Right(sSuffixNm, 1)
        Else
            lsValue = lsValue & Right(sFrstName, 1)
        End If

        lsValue = lsValue & Right(sMiddName, 1)
        lsValue = lsValue & Right(sLastName, 1)

        Return lsValue.ToUpper
endProc:
        Exit Function
    End Function

    Private Function RandInRange(ByVal fnMin As Integer, ByVal fnMax As Integer) As Integer
        If (fnMax < fnMin) Then Return 0

        Randomize()
        Return CInt(Int((fnMax - fnMin + 1) * Rnd() + fnMin))
    End Function

    Public Function Dec2Radix(ByVal TempDec As Long, ByVal Radix As Integer) As String
        Dim TNo As Integer
        Do
            TNo = CInt(TempDec - (Fix(TempDec / Radix) * Radix))
            If TNo > 9 Then
                Dec2Radix = Chr(55 + TNo) & Dec2Radix
            Else
                Dec2Radix = TNo & Dec2Radix
            End If
            TempDec = CLng(Fix(TempDec / Radix))
        Loop Until (TempDec = 0)
    End Function

    Public Function Radix2Dec(ByVal StrNum As String, ByVal Radix As Integer) As Long
        Dim lngOut As Long
        Dim i As Integer
        Dim c As Integer

        For i = 1 To Len(StrNum)
            c = Asc(UCase(Mid(StrNum, i, 1)))
            Select Case c
                Case 48 To 57
                    lngOut = CLng(lngOut + ((c - 48) * Radix ^ (Len(StrNum) - i)))
                Case Else
                    lngOut = CLng(lngOut + ((c - 55) * Radix ^ (Len(StrNum) - i)))
            End Select
        Next i
        Radix2Dec = lngOut
    End Function

    Private Function GetSplitedName(ByVal lsName As String) As Object
        Dim lasName() As String
        Dim lsLName As String, lsFName As String, lsMName As String
        Dim lnCtr As Integer

        lsLName = ""
        lsFName = ""
        lsMName = ""

        If lsName = "" Then
            ReDim lasName(0)
            lsLName = ""
        Else
            lasName = Split(Trim(lsName), ",")
            lsLName = Trim(UCase(Left(lasName(0), 1)) & Mid(lasName(0), 2))
        End If

        If UBound(lasName) > 0 Then
            If Trim(lasName(1)) <> "" Then
                lasName = Split(Trim(lasName(1)), " ")
                lsFName = Trim(UCase(Left(lasName(0), 1)) & Mid(lasName(0), 2))
                If UBound(lasName) > 0 Then
                    lsMName = Trim(UCase(Left(lasName(UBound(lasName)), 1)) & Mid(lasName(UBound(lasName)), 2))
                    For lnCtr = 1 To UBound(lasName) - 1
                        lsFName = lsFName & " " & Trim(UCase(Left(lasName(lnCtr), 1)) & Mid(lasName(lnCtr), 2))
                    Next
                End If
            End If
        End If

        ReDim lasName(2)
        lasName(0) = lsLName
        lasName(1) = lsFName
        lasName(2) = lsMName

        Return lasName
    End Function
End Module
