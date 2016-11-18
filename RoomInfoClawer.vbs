SelfPath = CreateObject("Scripting.FileSystemObject").GetFile(Wscript.ScriptFullName).ParentFolder.Path & "\"
lstPath = SelfPath & "room.txt"
'Set file = CreateObject("Scripting.FileSystemObject").OpenTextFile(lstPath, 2, True)
Call FileWriter(lstPath, "", False)

'"http://202.119.250.127:7555/ST.aspx"
Set Shell = CreateObject("Shell.Application")
For Each win In Shell.Windows
    If InStr(win.LocationURL, "202.119.250.127") > 0 Then
        Set IE = win
    End If
Next
Set Shell = Nothing
'msgbox IE.document.getElementsByTagName("title")(0).text
'msgbox IE.document.getElementById("BianHao").innerText
'Call DoSelect("DDLLouCeng",1)
'msgbox DoSelect("DDLFangJian",1)

ub1 = IE.document.getElementById("DDLSheQu").options.length - 1
For i = 0 To ub1
    DDLSheQuName = DoSelect("DDLSheQu", ub1 - i) & "#"

    ub2 = IE.document.getElementById("DDLLouYu").options.length - 1
    For j = 0 To ub2
        DDLLouYuName = DDLSheQuName & DoSelect("DDLLouYu", ub2 - j) & "#"

        strLouYuFull = ""

        ub3 = IE.document.getElementById("DDLDanYuan").options.length - 1
        For k = 0 To ub3
            DDLDanYuanName = DDLLouYuName & DoSelect("DDLDanYuan", ub3 - k) & "#"

            ub4 = IE.document.getElementById("DDLLouCeng").options.length - 1
            For l = 0 To ub4
                DDLLouCengName = DDLDanYuanName &  DoSelect("DDLLouCeng", ub4 - l) & "#"

                ub5 = IE.document.getElementById("DDLFangJian").options.length - 1
                For m = 0 To ub5
                    DDLFangJianName = DDLLouCengName &  DoSelect("DDLFangJian", ub5 - m)

                    'msgbox Trim(IE.document.getElementById("BianHao").innerText) & VbTab & DDLFangJianName
                    'file.WriteLine Trim(IE.document.getElementById("BianHao").innerText) & VbTab & DDLFangJianName
                    'Call FileWriter(lstPath, Trim(IE.document.getElementById("BianHao").innerText) & VbTab & DDLFangJianName, True)
                    strLouYuFull = strLouYuFull & Trim(IE.document.getElementById("BianHao").innerText) & VbTab & DDLFangJianName & VbCrLf

                Next
            Next
        Next

        Call FileWriter(lstPath, strLouYuFull, True)

    Next
Next

'file.WriteLine "Done"
'file.Close
'Set file = Nothing
msgbox "Done"

Function DoSelect(selectId,selectIndex)
    strBefore = IE.document.getElementById("BianHao").innerText
    IE.document.getElementById(selectId).options(selectIndex).selected = true
    IE.document.getElementById(selectId).onchange()
    'If (IE.document.getElementById(selectId).options.length = 1) Then
        'WScript.Sleep 1000
    'Else
    count = 0
    Do
        WScript.Sleep 10
        count = count + 1
    Loop While (strBefore = IE.document.getElementById("BianHao").innerText) Or (count > 500)
    'End If
    DoSelect = Trim(IE.document.getElementById(selectId)(IE.document.getElementById(selectId).selectedIndex).text)
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
