On Error Resume Next
SelfPath = CreateObject("Scripting.FileSystemObject").GetFile(Wscript.ScriptFullName).ParentFolder.Path & "\"
If (InStr(SelfPath, "Dell") > 0) Then
    RootPath = SelfPath & "History\"
    isCors = True
Else
    RootPath = "E:\Elec\"
    isCors = False
End If
NowHour = FormatDT(Now, "hh")
bHour = 8
eHour = 23
HostUrl = "http://202.119.250.127:7555/Login.aspx"
Call MakeUrlT(HostUrl, UrlT1, UrlT2)
UrlT2T = "&BtnLogin=" & URLEncode("�� ¼") & "&TxtPwd="
If (UrlT2 <> UrlT2T) Then
    Call WeChatAlert("��վ����!")
    Call FileWriter(RootPath & "Err1.txt", LongNow() & VbTab & GetPage(HostUrl), True)
    WScript.Quit
End If
lstPath = RootPath & "room.lst"
lstTxt = FileReader(lstPath)
msgs = ""
If (Trim(lstTxt) <> "") Then
    lstArray = Split(lstTxt, VbCrlf)
    For Each tmpLine In lstArray
        If (InStr(tmpLine, VbTab) > 0 And Left(tmpLine, 1) <> "'") Then
            roomInfo = Split(tmpLine, VbTab)
            'msgbox Trim(roomInfo(0))
            'msgbox UBound(roomInfo)
            msgs = msgs & CheckElec(Trim(roomInfo(0)), Trim(roomInfo(1)), Trim(roomInfo(2)))
        End If
    Next
End If

If (InStr(msgs, "������") > 0 And isCors = False) Then
    Msgbox msgs
End If

Sub MakeUrlT(HostUrl, UrlT1, UrlT2)
    strHtml = GetPage(HostUrl)
    strVIEWSTATE = URLEncode(GetElementValueByIdFromHtml("__VIEWSTATE", strHtml))
    strVIEWSTATEGENERATOR = URLEncode(GetElementValueByIdFromHtml("__VIEWSTATEGENERATOR", strHtml))
    strEVENTVALIDATION = URLEncode(GetElementValueByIdFromHtml("__EVENTVALIDATION", strHtml))
    strBtnLogin = URLEncode(GetElementValueByIdFromHtml("BtnLogin", strHtml))
    UrlT1 = HostUrl + "?__VIEWSTATE=" + strVIEWSTATE + "&__VIEWSTATEGENERATOR=" + strVIEWSTATEGENERATOR + "&__EVENTVALIDATION=" + strEVENTVALIDATION + "&TxtName="
    UrlT2 = "&BtnLogin=" & strBtnLogin & "&TxtPwd="
End Sub

Function WeChatAlert(msgs)
    url = "http://sc.ftqq.com/SCU2447T4ca73300c167a42a530ade37c0ea7ddb57e66eba92992.send?text=Elec&desp=" & msgs
    strHtml = GetPage(url)
    If (InStr(strHtml, "success") > 0) Then
        WeChatAlert = True
    Else
        WeChatAlert = False
    End If
End Function

Function FileWriter(fName, msg, bAppending)
    fApd = 8
    If (bAppending = False) Then
        fApd = 2
    End If
    Set file = CreateObject("Scripting.FileSystemObject").OpenTextFile(fName, fApd, True)
    file.WriteLine msg
    file.Close
    Set file = Nothing
End Function

Function FileReader(fName)
    fileTxt = ""
    Set fso = CreateObject("Scripting.FileSystemObject")
    If (fso.FileExists(fName)) Then
        Set file = fso.OpenTextFile(fName, 1, False)
        fileTxt = file.ReadAll
        file.Close
        Set file = Nothing
        End If
    Set fso = Nothing
    FileReader = fileTxt
End Function

Function CheckElec(room, psw, name)
    rst = ""
    difElec = 0

    url = UrlT1 & room & UrlT2 & psw
    strHtml = GetPage(url)

    If (InStr(strHtml, "���ϵ��������ο�") > 0) Then
        'elec = Trim(GetElementInnerTextByIdFromHtml("LabElec", strHtml))
        elec = GetElementInnerTextByIdFromHtml("LabElec", strHtml)

        If (isNumeric(elec)) Then
            rst = rst & ElecTest(strHtml)

            If (name = "") Then
                name = Trim(GetElementInnerTextByIdFromHtml("LabCommunity", strHtml)) + "_" + Trim(GetElementInnerTextByIdFromHtml("LabHouseNum", strHtml))
            End If

            txtPath = RootPath & room & "_" & name & ".txt"
            Set fso = CreateObject("Scripting.FileSystemObject")
            strLine = ""
            strTxt = FileReader(txtPath)
            If (Trim(strTxt) <> "") Then
                For Each tmpLine In Split(strTxt, VbCrlf)
                    If (InStr(tmpLine, VbTab) > 0) Then
                        strLine = tmpLine
                    End If
                Next

                '���ļ���object.OpenTextFile(filename[, iomode{1=ForReading;2=ForWriting3=ForAppending}[, create[, format{null=ansi}]]])
                Set file = fso.OpenTextFile(txtPath, 8, False)
            Else
                '�����ļ���object.CreateTextFile(filename,[ overwrite,[ unicode{True=UTF8}]])
                Set file = fso.CreateTextFile(txtPath, False, False)
            End If

            If (Trim(strLine) <> "") Then
                lastElec = Split(strLine, VbTab)(1)
                difElec = lastElec - elec
                'д������
                'If (InStr(strLine,Split(LongNow(), " ")(0)) = 0 And (lastElec <> elec Or elec = 0)) Then
                'If (lastElec <> elec Or (elec = 0 And InStr(strLine, Split(LongNow(), " ")(0)) = 0)) Then
                '(�����仯)��(����û�������³���ʱ���������¼ʱ��)
                If (difElec <> 0 Or (difElec = 0 And InStr(strLine, FormatDT(GetLastTimeFromHtml(strHtml), "yyyy-mm-dd")) = 0)) Then
                    'д���ļ����ݣ������ַ�����write(x)д��x���ַ���writelineд�뻻�У�writeblanklines(n)д��n������
                    file.WriteLine LongNow() & VbTab & elec
                End If

                '�ж��쳣
                If (difElec > 30 Or difElec < -500 Or elec > 1000) Then
                    rst = rst & VbTab & "�����쳣��" & VbTab & elec
                    Call FileWriter(RootPath & "ErrUrl.txt", LongNow() & VbTab & url, True)
                    Call FileWriter(RootPath & "Err3.txt", LongNow() & VbTab & strHtml, True)
		End If
            Else
                difElec = 1
                file.WriteLine LongNow() & VbTab & elec
            End If

            '���ļ�
            file.Close
            Set file = Nothing
            Set fso = Nothing

            '�жϵ�����
            If (elec < 20 And difElec <> 0) Then
                rst = rst & VbTab & "�����ͣ�" & VbTab & elec
            End If
        Else
            rst = rst & VbTab & "δ��ȡ��������"
        End If
    ElseIf (InStr(strHtml, "�û������������") > 0) Then
        rst = rst & VbTab & name & "�������" & room & "_" & psw
        Call FileWriter(RootPath & "ErrPsw.txt", LongNow() & VbTab & room & VbTab & psw & VbTab & name , True)
    Else
        rst = rst & VbTab & "��¼ʧ�ܣ�"
        Call FileWriter(RootPath & "ErrUrl.txt", LongNow() & VbTab & url, True)
        Call FileWriter(RootPath & "Err2.txt", LongNow() & VbTab & strHtml, True)
    End If
    If (rst = "") Then
    	CheckElec = rst
    Else
        Call WeChatAlert(name & rst)
        CheckElec = room & "_" & name & rst & chr(10)
    End If
End Function

Function ElecTest(strHtml)
    Set html = CreateObject("HtmlFile")
    html.designMode = "on" ' �����༭ģʽ
    html.write strHtml ' д������
    elec1 = html.getElementsByTagName("td")(11).innerText
    elec2 = html.getElementsByTagName("td")(21).innerText
    Set html = Nothing
    If (elec1 = elec2) Then
        test = ""
    Else
        test = elec1 & "��һ��" & elec2
    End If
    ElecTest = rst
End Function

Function LongNow()
    LongNow = FormatDT(Now, "yyyy-mm-dd hh:nn:ss")
End Function

Function FormatDT(aDateTime, interval)
    Set regEx = New RegExp
    regEx.IgnoreCase = True
    regEx.Global = True

    FormatDT = interval

    Const DBLPATTERNS = "m,d,h,n,s"
    Const PATTERNS = "yyyy,q,m,y,d,ww,w,h,n,s"
    For Each DblPtrn In Split(DBLPATTERNS, ",")
        regEx.Pattern = DblPtrn & DblPtrn
        sDatePart = DatePart(DblPtrn, aDateTime)
        If (sDatePart < 10) Then
            sDatePart = "0" & sDatePart
        End If
        FormatDT = regEx.Replace(FormatDT, sDatePart)
    Next
    For Each Ptrn In Split(PATTERNS, ",")
        regEx.Pattern = Ptrn
        sDatePart = DatePart(Ptrn, aDateTime)
        FormatDT = regEx.Replace(FormatDT, DatePart(Ptrn, aDateTime))
    Next
End Function

Function GetPage(url)
    'On Error Resume Next
    Set http = CreateObject("Msxml2.ServerXMLHTTP")
    http.open "GET", url, False
    count = 10
    Do
        http.send
        http.waitForResponse 200
        While http.readyState <> 4
            http.waitForResponse 200
        WEnd
        strHtml = http.responseText
        strTitle = ""
        strTitle = GetTitleFromHtml(strHtml)
        If (Trim(strTitle) = "Runtime Error") Then
            Call FileWriter(RootPath & "ErrPage.txt", LongNow() & VbTab & url & VbTab & strHtml, True)
            errFlag = True
        Else
            errFlag = False
        End If
        count = count - 1
        If (count < 0) Then
            errFlag = False
        End If
    Loop While errFlag
    Set http = Nothing

    GetPage = strHtml
End Function

Function GetElementInnerTextByIdFromHtml(strId, strHtml)
    Set html = CreateObject("HtmlFile")
    html.designMode = "on" ' �����༭ģʽ
    html.write strHtml ' д������
    strElement = ""
    strElement = html.getElementById(strId).innerText
    Set html = Nothing

    GetElementInnerTextByIdFromHtml = strElement
End Function

Function GetElementValueByIdFromHtml(strId, strHtml)
    Set html = CreateObject("HtmlFile")
    html.designMode = "on" ' �����༭ģʽ
    html.write strHtml ' д������
    strElement = ""
    strElement = html.getElementById(strId).Value
    Set html = Nothing

    GetElementValueByIdFromHtml = strElement
End Function

Function GetTitleFromHtml(strHtml)
    Set html = CreateObject("HtmlFile")
    html.designMode = "on" ' �����༭ģʽ
    html.write strHtml ' д������
    strElement = ""
    strElement = html.getElementsByTagName("title")(0).innerText
    Set html = Nothing

    GetTitleFromHtml = strElement
End Function

Function URLEncode(strURL)
    Dim I
    Dim tempStr
    For I = 1 To Len(strURL)
        If Asc(Mid(strURL, I, 1)) < 0 Then
            tempStr = "%" & Right(CStr(Hex(Asc(Mid(strURL, I, 1)))), 2)
            tempStr = "%" & Left(CStr(Hex(Asc(Mid(strURL, I, 1)))), Len(CStr(Hex(Asc(Mid(strURL, I, 1))))) - 2) & tempStr
            URLEncode = URLEncode & tempStr
        ElseIf (Asc(Mid(strURL, I, 1)) >= 65 And Asc(Mid(strURL, I, 1)) <= 90) Or (Asc(Mid(strURL, I, 1)) >= 97 And Asc(Mid(strURL, I, 1)) <= 122) Or (Asc(Mid(strURL, I, 1)) >= 48 And Asc(Mid(strURL, I, 1)) <= 57) Then
            URLEncode = URLEncode & Mid(strURL, I, 1)
        Else
            URLEncode = URLEncode & "%" & Hex(Asc(Mid(strURL, I, 1)))
        End If
    Next
End Function

Function GetLastTimeFromHtml(strHtml)
    Set html = CreateObject("HtmlFile")
    html.designMode = "on" ' �����༭ģʽ
    html.write strHtml ' д������
    Set elementArray = html.getElementsByTagName("td")
    strLastElement = ""
    For Each element In elementArray
        If (InStr(strLastElement, "ʣ�����") > 0) Then
            strLastElement = element.innerText
            Exit For
        End If
        strLastElement = element.innerText
    Next
    Set html = Nothing
    GetLastTimeFromHtml = strLastElement
End Function
