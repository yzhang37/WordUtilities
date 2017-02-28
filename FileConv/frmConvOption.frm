VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmConvOption 
   Caption         =   "文档转换实用工具"
   ClientHeight    =   10570
   ClientLeft      =   90
   ClientTop       =   410
   ClientWidth     =   16640
   OleObjectBlob   =   "frmConvOption.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "frmConvOption"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim folderType As String
Dim repeatError As String
Dim msgFinsihButton As String
Dim msgNextButton As String
Dim msgBackButton As String
Dim msgColuHeadFile As String
Dim msgColuHeadType As String
Dim msgColuHeadstatus As String
Dim msgColuHeadPlace As String
Dim msgColuConverter As String
Dim msgOutputType As String
Dim msgConvertStatus(100) As String
Dim msgConvertStatuWaiting As String
Dim msgStatusSucceed As String
Dim msgStatusCautious As String
Dim msgStatusFailed As String
Dim msgColuHeadInfo As String

Dim checkedSource As Long
Dim checkFlag As Boolean
Dim selectedTargetFormat As Long
Dim originalPathOutput As Boolean
Dim IsDeleteOriginal As Boolean

Dim saveFormatDesc(17) As String
Dim saveFormatSuffix(17) As String
Dim saveFormatCode(17) As Long
Dim mainFSO As FileSystemObject

Dim STRS(100) As String

Private Declare Function MultiByteToWideChar Lib "kernel32" (ByVal _
CodePage As Long, ByVal dwFlags As Long, ByRef lpMultiByteStr As Any, _
ByVal cchMultiByte As Long, ByVal lpWideCharStr As Long, ByVal _
cchWideChar As Long) As Long

Private Sub LoadStrings()
    folderType = "文件夹"
    repeatError = "文件内容重复"
    msgFinsihButton = "转换"
    msgNextButton = "下一步 >"
    msgBackButton = "< 上一步"
    msgColuHeadFile = "文件名称"
    msgColuHeadType = "类型"
    msgColuHeadPlace = "位置"
    msgOutputType = "输出文件类型:" + vbLf + "%1"
    
    Call Seperate("DOS 文本/DOS 文本(含换行符)/OpenDocument 文档/PDF/RTF 格式/Strict Open XML 文档/Word 97-2003 模板/Word 97-2003 文档/Word XML 2003 文档/Word XML 文档/Word 文档/XPS 文档/纯文本/纯文本 (含换行符)/启用宏的 Word 模板/启用宏的 Word 文档/筛选过的网页/网页/", _
              "/", saveFormatDesc)
    Call SeperateNumber("4,5,23,17,6,24,1,0,11,12,16,18,2,3,15,13,10,8", ",", saveFormatCode)
    Call Seperate(".txt/.txt/.odt/.pdf/.RTF/.docx/.dot/.doc/.XML/.XML/.docx/.xps/.txt/.txt/.dotm/.docm/.htm/.htm/", "/", saveFormatSuffix)

    msgConvertStatuWaiting = "等待转换"
    msgColuHeadInfo = "注意事项"
    msgColuHeadstatus = "状态"
    msgColuConverter = "转换器名称"
    msgStatusPassed = "目标位置无法创建。"
    
    msgStatusSucceed = "完成"
    msgStatusCautious = "注意"
    msgStatusFailed = "失败"
    
    STRS(0) = "目标存放路径不存在"
    STRS(1) = "你输入的目标位置 '%1' 不存在。"
    STRS(2) = "请按'取消'重新检查，然后再试一次。"
    STRS(3) = "确定"
    STRS(4) = "取消"
    STRS(5) = "Microsoft Word 实用工具"
    STRS(6) = "文档转换实用工具"
    STRS(7) = "批量文档格式转换"
    STRS(8) = "批量转换 Microsoft Word 文档格式"
    STRS(9) = "欢迎！这是一款小型迷你的，内嵌在 Word 程序中的批量 Word 转换工具。" & _
              "你可以用它来将你磁盘中大量的陈旧的格式批量自动转换成为最新的格式。" & vbLf & vbLf & _
              "点击""下一页""既可以开始转换。"
    STRS(10) = "下一次跳过这个页面"
    STRS(11) = "选择要转换的文档"
    STRS(12) = "可以添加多个文件或者多个文件夹。"
    STRS(13) = "浏览..."
    STRS(14) = "浏览文件夹..."
    STRS(15) = "转换后删除原来的文件"
    STRS(16) = "转换多个文件的格式"
    STRS(17) = "选择目标文档的格式，然后点击下一步开始转换。"
    STRS(18) = "选择存放转换文档的目标文件夹"
    STRS(19) = "存放输出文件到原文件夹中"
    STRS(20) = "或输出到以下位置"
    STRS(21) = "正在转换文档"
    STRS(22) = "批量转换文档格式中..."
    STRS(23) = "转换结果"
    STRS(24) = "以下是转换所有文档的结果。请检查'状态'一栏以确保全部转换。"
    STRS(25) = "完成"
    STRS(26) = "保存结果..."
    
    msgConvertStatus(StatusOK) = "转换完成，在 '%1' "
    msgConvertStatus(StatusOKAndDeletefailed) = "文件已转换，但是删除原文件 '%2' 失败，可能是文件被占用，或者没有权限访问该文件。"
    msgConvertStatus(StatusRenamed) = "完成并已更名为 '%1' 。"
    msgConvertStatus(StatusRenamedAndDeletefailed) = "完成并已更名为 '%1' ，但是删除原文件 '%2' 失败，可能是文件被占用，或者没有权限访问该文件。"
    msgConvertStatus(StatusReplaced) = "完成并已替换重名文件 '%1' 。"
    msgConvertStatus(StatusReplacedAndDeletefailed) = "完成并已替换重名文件 '%1' ，但是删除原文件 '%2' 失败，可能是文件被占用，或者没有权限访问该文件。"
    
    msgConvertStatus(StatusFailedforReplace) = "因为无法替换已存在的文件 '%1' 而跳过。"
    msgConvertStatus(StatusFailedforRename) = "因为存在重名文件 '%1' 而跳过。"
    msgConvertStatus(StatusFailedforCancel) = "被用户取消转换。"
    msgConvertStatus(StatusFailedforOpen) = "打开原文件 '%2' 时失败，可能是文件被占用，磁盘损坏，或者没有权限访问该文件。"
    msgConvertStatus(StatusFailedforSave) = "转换时无法生成目标文件 '%2' ，可能没有权限存取该文件或因磁盘损坏无法存取。"
    
    STRS(27) = "本次文件转换结果如下："
    STRS(28) = "共计转换 %1 个文件，其中 %2 个成功， %3 个失败。"
    STRS(29) = "转换成功率为 %1 %。"
    STRS(29) = "转换成功率约 %1 %"
    STRS(31) = "以下文件因为错误而转换失败，请重新核对："
    STRS(32) = "没有文件因为错误而转换失败。Nice!"
    STRS(33) = "文件名称"
    STRS(34) = "失败原因"
    STRS(35) = "以下文件转换过程中系统自动做了一些变动，请核对："
    STRS(36) = "没有文件有需要注意的事项。"
    STRS(37) = "系统改变"
    STRS(38) = "以下文件转换成功完成："
    STRS(39) = "转换结果"
    STRS(40) = "本次没有转换任何文件。"
    InitUI
End Sub

'Private Sub cmdSaveToFile_Click()
'    Dim msg As String
'    WriteToFile (msg)
'    MsgBox msg
'End Sub

'Private Sub WriteToFile(buffer As String)
'    buffer = ""
'    Link buffer, STRS(27)
'
'    Dim s_succ As String, s_caut As String, s_fail As String
'    Dim t_succ As Long, t_caut As Long, t_fail As Long
'    Dim i As Integer
'    For i = 1 To lstWaiting.ListItems.Count
'        With lstWaiting.ListItems(i)
'            Select Case .SubItems(1)
'            Case msgStatusSucceed
'                If t_succ = 0 Then Link s_succ, STRS(38)
'                t_succ = t_succ + 1
'
'            Case msgStatusCautious
'                If t_caut = 0 Then Link s_caut, STRS(35)
'                t_caut = t_caut + 1
'
'            Case msgStatusFailed
'                If t_fail = 0 Then Link s_fail, STRS(31)
'                t_fail = t_fail + 1
'
'            End Select
'        End With
'    Next i
'
'
'    Link buffer, Replace(Replace(Replace(STRS(28), "%1", "1"), "%2", "1"), "%3", "0")
'End Sub

Private Sub Link(StringA As String, StringB As String)
    StringA = StringA & StringB & vbLf
End Sub

Private Sub cbxDontDisplayNext_Click()
    Dim dont As Long
    Call SaveSetting("OfficeUtilities", "DocConv", _
        "DoNotShowFirstPage", Trim(Str(IIf(cbxDontDisplayNext.value, 1, 0))))
End Sub

Private Sub cbxOriginalFolder_Click()
    tbxDestPath.Enabled = Not cbxOriginalFolder.value
    cmdDestPathSelector.Enabled = tbxDestPath.Enabled
    If Not tbxDestPath.Enabled Then
        tbxDestPath.BackColor = &H8000000F
    Else
        tbxDestPath.BackColor = &H80000005
    End If
End Sub

Private Sub cbxRemoveOriginalFiles_Click()
    IsDeleteOriginal = cbxRemoveOriginalFiles.value
End Sub

Private Sub cmdBrowse_Click()
    On Error Resume Next
    Dim i As Integer
    Dim sF As String
    Dim sFilInfo As SHFILEINFO
    cmdSelect.value = False
    Set fDlg = Application.FileDialog(msoFileDialogOpen)
    lstboxSource.Sorted = False
    lstboxSource.Checkboxes = True
    With fDlg
        For i = 1 To .Filters.Count
            If InStr(1, fDlg.Filters(i).Extensions, "*.*") > 0 Then
                fDlg.Filters.Delete i
                Exit For
            End If
        Next
        .FilterIndex = 0
        .AllowMultiSelect = True
        rValue = .Show()
        If rValue = True Then
            For Each sFile In .SelectedItems
                i = lstboxSource.ListItems.Count + 1
                If Err.Number = 35602 Then
                    
                Else
                    sF = GetFileName(sFile)
                    lstboxSource.ListItems.Add i, sFile, sF
                    SHGetFileInfo sFile, 0, sFilInfo, Len(sFilInfo), SHGFI_TYPENAME Or SHGFI_DISPLAYNAME
                    lstboxSource.ListItems.Item(i).SubItems(1) = sFilInfo.szTypeName
                    lstboxSource.ListItems.Item(i).SubItems(2) = GetPath(sFile)
                End If
            Next
        End If
    End With
    lstboxSource.Sorted = True
    lstboxSource.Checkboxes = False
End Sub

Private Sub cmdBrowseFolder_Click()
    On Error Resume Next
    Dim i As Integer
    cmdSelect.value = False
    Set fDlg = Application.FileDialog(msoFileDialogFolderPicker)
    lstboxSource.Sorted = False
    With fDlg
        rValue = .Show()
        If rValue = True Then
            For Each sFile In .SelectedItems
                i = lstboxSource.ListItems.Count + 1
                lstboxSource.ListItems.Add i, sFile, GetFileName(sFile)
                lstboxSource.ListItems.Item(i).SubItems(1) = folderType
                lstboxSource.ListItems.Item(i).SubItems(2) = GetPath(sFile)
            Next
        End If
    End With
    lstboxSource.Sorted = True
End Sub

Private Sub cmdCancel_Click()
    End
End Sub

Private Sub cmdDeleteSource_Click()
    If checkedSource > 0 Then
        Dim i As Long
        i = 1
        While i <= lstboxSource.ListItems.Count
            If lstboxSource.ListItems(i).checked = True Then
                lstboxSource.ListItems.Remove i
            Else
                i = i + 1
            End If
        Wend
        checkedSource = 0
        cmdSelectAll_Click
    End If
    lstboxSource.SetFocus
End Sub

Private Sub cmdDestPathSelector_Click()
    Set foldDlg = Application.FileDialog(msoFileDialogFolderPicker)
    With foldDlg
        .InitialFileName = tbxDestPath.Text
        If .Show Then
            tbxDestPath.Text = .SelectedItems(1)
        End If
    End With
End Sub

Private Sub cmdNext_Click()
    Navigate
End Sub

Private Sub cmdNext_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    cmdNext_Click
End Sub

Private Sub cmdBack_Click()
    Navigate -2
End Sub

Private Sub cmdBack_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    cmdBack_Click
End Sub

Private Sub cmdSelect_Click()
    checkFlag = False
    If Not cmdSelect.value Then
        cmdSelectAll.value = False
        cmdSelectAll_Click
    End If
    cmdDeleteSource.Visible = cmdSelect.value
    cmdSelectAll.Visible = cmdSelect.value
    lstboxSource.Checkboxes = cmdSelect.value
    lstboxSource.Refresh
End Sub

Private Sub cmdSelectAll_Click()
    If checkFlag Then Exit Sub
    lstboxSource.SetFocus
    If lstboxSource.ListItems.Count = 0 Then
        checkFlag = True
        cmdSelectAll.value = False
        checkFlag = False
        Exit Sub
    End If
    If cmdSelectAll.value = True Then
        For Each Item In lstboxSource.ListItems
            Item.checked = True
        Next
        checkedSource = lstboxSource.ListItems.Count
    Else
        For Each Item In lstboxSource.ListItems
            Item.checked = False
        Next
        checkedSource = 0
    End If
End Sub


Private Sub lstboxSource_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    If lstboxSource.SortKey + 1 <> ColumnHeader.index Then
        lstboxSource.SortKey = ColumnHeader.index - 1
        lstboxSource.SortOrder = lvwAscending
    Else
        lstboxSource.SortOrder = IIf(lstboxSource.SortOrder = lvwAscending, ListSortOrderConstants.lvwDescending, ListSortOrderConstants.lvwAscending)
    End If
End Sub

Private Sub lstboxSource_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    If Item.checked Then
        checkedSource = checkedSource + 1
    Else
        checkedSource = checkedSource - 1
    End If
    checkFlag = True
    If checkedSource = lstboxSource.ListItems.Count Then
        cmdSelectAll.value = True
    Else
        cmdSelectAll.value = False
    End If
    checkFlag = False
End Sub

Private Sub lstFileType_ItemClick(ByVal Item As MSComctlLib.ListItem)
    lblOutputType.Caption = Replace(msgOutputType, "%1", Item)
    selectedTargetFormat = Item.index - 1
End Sub


Private Sub lstWaiting_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    If lstWaiting.Sorted = False Then Exit Sub
    If lstWaiting.SortKey + 1 <> ColumnHeader.index Then
        lstWaiting.SortKey = ColumnHeader.index - 1
        lstWaiting.SortOrder = lvwAscending
    Else
        lstWaiting.SortOrder = IIf(lstWaiting.SortOrder = lvwAscending, ListSortOrderConstants.lvwDescending, ListSortOrderConstants.lvwAscending)
    End If
End Sub

Private Sub UserForm_Initialize()
    Set mainFSO = New FileSystemObject
    LoadStrings
    ChDir Environ("HOMEDRIVE") + Environ("HOMEPATH")
    cbxOriginalFolder_Click
    lstboxSource.ColumnHeaders.Add , "name", msgColuHeadFile
    lstboxSource.ColumnHeaders.Add , "type", msgColuHeadType
    lstboxSource.ColumnHeaders.Add , "folder", msgColuHeadPlace
    lstboxSource.SortKey = 1
    lstFileType.ColumnHeaders.Add , "desc", msgColuConverter
    lstboxSource.ColumnHeaders.Item(1).Width = lstboxSource.Width * 5 / 10
    lstboxSource.ColumnHeaders.Item(2).Width = lstboxSource.Width * 3 / 10
    lstboxSource.ColumnHeaders.Item(3).Width = lstboxSource.Width * 2 / 10
    lstFileType.ColumnHeaders.Item(1).Width = lstFileType.Width * 9 / 10
    
    With lstWaiting
        .ColumnHeaders.Add , "name", msgColuHeadFile
        .ColumnHeaders.Add , "status", msgColuHeadstatus
        .ColumnHeaders.Add , "info", msgColuHeadInfo
        .ColumnHeaders.Item(1).Width = .Width * 4.4 / 10
        .ColumnHeaders.Item(2).Width = .Width * 1.2 / 10
        .ColumnHeaders.Item(3).Width = .Width * 4.4 / 10
    End With
    
    Set fDlg = Application.FileDialog(msoFileDialogSaveAs)
    
    Dim i As Integer
    For i = 0 To 17
        lstFileType.ListItems.Add i + 1, , saveFormatDesc(i)
    Next
    
    IsDeleteOriginal = False
    cbxRemoveOriginalFiles.value = IsDeleteOriginal
    selectedTargetFormat = 10
    Set lstFileType.SelectedItem = lstFileType.ListItems(selectedTargetFormat + 1)
    
    lstFileType_ItemClick lstFileType.SelectedItem
    
    tbxDestPath.Text = CurDir

    cmdSelect_Click
    
    Me.Width = 405
    Me.Height = 306
    
    Dim dontshowfirst As Long
    dontshowfirst = Val(GetSetting("OfficeUtilities", "DocConv", "DoNotShowFirstPage", "0"))
    If dontshowfirst <> 0 Then
        Navigate 1
        cbxDontDisplayNext.value = True
    Else
        Navigate 0
    End If
End Sub

Private Sub Navigate(Optional index As Integer = -1)
    Static current As Integer
    Dim raw As Integer
    raw = current
    cmdSelect.value = False
    Select Case index
    Case -1
        current = current + 1
    Case -2
        current = current - 1
    Case Else
        current = index
    End Select
    If current < 0 Then
        current = raw
    End If
Navigate:
    cmdBack.Enabled = (current > 0) And (current < 3)
    panel1.Visible = False
    panel2.Visible = False
    panel3.Visible = False
    panel4.Visible = False
    cmdNext.Enabled = (current <> 3)
    With panel1
        If current < 2 Then
            cmdNext.Caption = msgNextButton
        Else
            cmdNext.Caption = msgFinsihButton
        End If
        Select Case current
        Case 0
            panel1.Visible = True
            cbxDontDisplayNext.SetFocus
        Case 1
            panel2.Move .Left, .Top, .Width, .Height
            panel2.Visible = True
            cmdBrowse.SetFocus
        Case 2
            panel3.Move .Left, .Top, .Width, .Height
            panel3.Visible = True
            lstFileType.SetFocus
        Case 3
            panel4.Move .Left, .Top, .Width, .Height
            panel4.Visible = True
            'it looks like busy working
            panel4.MousePointer = fmMousePointerHourGlass
            lstWaiting.ListItems.Clear
            
            If Not cbxOriginalFolder.value Then
                If Not mainFSO.FolderExists(tbxDestPath.Text) Then
                    Dim chkStatus As Boolean
                    QueryBox chkStatus, STRS(0), Replace(STRS(1), "%1", tbxDestPath.Text), _
                             STRS(4), , , 1, False, 1, STRS(2)
                    current = 2
                    GoTo Navigate
                End If
            End If
            
            Dim i As Integer
            For i = 1 To lstboxSource.ListItems.Count
                With lstboxSource.ListItems(i)
                    If .SubItems(1) = folderType Then
                        Call AddFolderName(.Key)
                    Else
                        Call AddFiles(.Key)
                    End If
                End With
            Next i
            StartConverting
            current = 4
            GoTo Navigate
        Case 4
            panel4.Move .Left, .Top, .Width, .Height
            panel4.Visible = True
            lstWaiting.Sorted = True
            lstWaiting.SortKey = 1
            progStatus.Visible = False
            Label7.Caption = STRS(23)
            Label8.Caption = STRS(24)
            cmdCancel.Enabled = False
            cmdNext.Caption = STRS(25)
            cmdNext.Enabled = True
            'cmdSaveToFile.Visible = True
            panel4.MousePointer = fmMousePointerArrow
        Case 5
            End
        End Select
    End With
End Sub

Private Sub StartConverting()
    Dim i As Long
    Dim myForce As FileForceSaveConstants
    Dim result As ConvertResultConstants
    
    myForce = FileForceSaveConstants.None
    
    Dim fold As Folder
    Dim targetDir As String
    
    '用于存储哪些错误提示框不要再显示了
    Dim DontShowAgain As Long
    DontShowAgain = 0
    
    If Not cbxOriginalFolder.value Then
        Set fold = mainFSO.GetFolder(tbxDestPath.Text)
        targetDir = fold.path
    End If
    For i = 1 To lstWaiting.ListItems.Count
        Dim newFileName As String
        Dim pureFileName As String
        
        With lstWaiting.ListItems(i)
            pureFileName = RemoveSuffix(GetFileName(.Key)) + saveFormatSuffix(selectedTargetFormat)
            If cbxOriginalFolder.value Then
                newFileName = GetPath(.Key) + "\" + pureFileName
            Else
                newFileName = targetDir + "\" + pureFileName
            End If
            
            result = DocConvert(mainFSO, ByVal .Key, _
                                newFileName, saveFormatCode(selectedTargetFormat), DontShowAgain, IsDeleteOriginal, myForce)
            If result <= StatusPerfect Then
                .SubItems(1) = msgStatusSucceed
            ElseIf result <= StatusNormal Then
                .SubItems(1) = msgStatusCautious
            Else
                .SubItems(1) = msgStatusFailed
            End If
            .SubItems(2) = Replace(Replace(msgConvertStatus(result), "%1", newFileName), "%2", .Key)
        End With
        progStatus.value = Int(progStatus.Min + (progStatus.Max - progStatus.Min) / _
                           lstWaiting.ListItems.Count * i)
        DoEvents
    Next i
End Sub

Private Sub Seperate(data As String, seperateSymbol As String, ByRef storage() As String)
    Dim i As Long, n As Long
    Dim sep As String
    i = 0
    n = 1
    nLen = InStr(n, data, seperateSymbol)
    Do While nLen > 0
        sep = Mid(data, n, nLen - n)
        storage(i) = sep
        n = nLen + 1
        nLen = InStr(n, data, seperateSymbol)
        i = i + 1
    Loop
End Sub

Private Sub SeperateNumber(data As String, seperateSymbol As String, ByRef storage() As Long)
    Dim i As Long, n As Long
    Dim sep As String
    i = 0
    n = 1
    nLen = InStr(n, data, seperateSymbol)
    Do While nLen > 0
        sep = Mid(data, n, nLen - n)
        storage(i) = Val(sep)
        n = nLen + 1
        nLen = InStr(n, data, seperateSymbol)
        i = i + 1
    Loop
End Sub

Private Sub AddFiles(FileName As String)
    On Error Resume Next
    lstWaiting.ListItems.Add , FileName, FileName
    lstWaiting.ListItems(lstWaiting.ListItems.Count).SubItems(1) = msgConvertStatuWaiting
End Sub

Private Sub AddFolder(ByRef fld As Folder)
    On Error Resume Next
    Dim fil As File
    Dim need As Boolean
    Dim subfold As Folder
    Dim thisSuffix As String
    For Each fil In fld.Files
        need = False
        thisSuffix = "." + LCase(GetSuffix(fil.Name))
        For i = 0 To 17
            If saveFormatSuffix(i) = thisSuffix Then
                need = True
                Exit For
            End If
        Next i
        If need Then
            Call AddFiles(fil.ParentFolder + "\" + fil.Name)
        End If
        DoEvents
    Next
    For Each subfold In fld.SubFolders
        Call AddFolder(subfold)
        DoEvents
    Next
End Sub

Private Sub AddFolderName(FolderStr As String)
    Dim fld As Folder
    Set fld = mainFSO.GetFolder(FolderStr)
    Call AddFolder(fld)
End Sub

Private Sub InitUI()
    Me.Caption = STRS(6)
    label0.Caption = STRS(5)
    Label1.Caption = STRS(7)
    Label2.Caption = STRS(8)
    lblWelcome.Caption = STRS(9)
    cbxDontDisplayNext.Caption = STRS(10)
    Label3.Caption = STRS(11)
    Label4.Caption = STRS(12)
    cmdBrowse.Caption = STRS(13)
    cmdBrowseFolder.Caption = STRS(14)
    cmdDestPathSelector.Caption = STRS(14)
    cbxRemoveOriginalFiles.Caption = STRS(15)
    Label5.Caption = STRS(16)
    Label6.Caption = STRS(17)
    frmSelectDir.Caption = STRS(18)
    cbxOriginalFolder.Caption = STRS(19)
    Label9.Caption = STRS(20)
    Label7.Caption = STRS(21)
    Label8.Caption = STRS(22)
    cmdSaveToFile.Caption = STRS(26)
    cmdCancel.Caption = STRS(4)
    cmdBack.Caption = msgBackButton
    cmdNext.Caption = msgNextButton
End Sub
