VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmConvOption 
   Caption         =   "�ĵ�ת��ʵ�ù���"
   ClientHeight    =   10570
   ClientLeft      =   90
   ClientTop       =   410
   ClientWidth     =   16640
   OleObjectBlob   =   "frmConvOption.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  '����������
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
    folderType = "�ļ���"
    repeatError = "�ļ������ظ�"
    msgFinsihButton = "ת��"
    msgNextButton = "��һ�� >"
    msgBackButton = "< ��һ��"
    msgColuHeadFile = "�ļ�����"
    msgColuHeadType = "����"
    msgColuHeadPlace = "λ��"
    msgOutputType = "����ļ�����:" + vbLf + "%1"
    
    Call Seperate("DOS �ı�/DOS �ı�(�����з�)/OpenDocument �ĵ�/PDF/RTF ��ʽ/Strict Open XML �ĵ�/Word 97-2003 ģ��/Word 97-2003 �ĵ�/Word XML 2003 �ĵ�/Word XML �ĵ�/Word �ĵ�/XPS �ĵ�/���ı�/���ı� (�����з�)/���ú�� Word ģ��/���ú�� Word �ĵ�/ɸѡ������ҳ/��ҳ/", _
              "/", saveFormatDesc)
    Call SeperateNumber("4,5,23,17,6,24,1,0,11,12,16,18,2,3,15,13,10,8", ",", saveFormatCode)
    Call Seperate(".txt/.txt/.odt/.pdf/.RTF/.docx/.dot/.doc/.XML/.XML/.docx/.xps/.txt/.txt/.dotm/.docm/.htm/.htm/", "/", saveFormatSuffix)

    msgConvertStatuWaiting = "�ȴ�ת��"
    msgColuHeadInfo = "ע������"
    msgColuHeadstatus = "״̬"
    msgColuConverter = "ת��������"
    msgStatusPassed = "Ŀ��λ���޷�������"
    
    msgStatusSucceed = "���"
    msgStatusCautious = "ע��"
    msgStatusFailed = "ʧ��"
    
    STRS(0) = "Ŀ����·��������"
    STRS(1) = "�������Ŀ��λ�� '%1' �����ڡ�"
    STRS(2) = "�밴'ȡ��'���¼�飬Ȼ������һ�Ρ�"
    STRS(3) = "ȷ��"
    STRS(4) = "ȡ��"
    STRS(5) = "Microsoft Word ʵ�ù���"
    STRS(6) = "�ĵ�ת��ʵ�ù���"
    STRS(7) = "�����ĵ���ʽת��"
    STRS(8) = "����ת�� Microsoft Word �ĵ���ʽ"
    STRS(9) = "��ӭ������һ��С������ģ���Ƕ�� Word �����е����� Word ת�����ߡ�" & _
              "�������������������д����ĳ¾ɵĸ�ʽ�����Զ�ת����Ϊ���µĸ�ʽ��" & vbLf & vbLf & _
              "���""��һҳ""�ȿ��Կ�ʼת����"
    STRS(10) = "��һ���������ҳ��"
    STRS(11) = "ѡ��Ҫת�����ĵ�"
    STRS(12) = "������Ӷ���ļ����߶���ļ��С�"
    STRS(13) = "���..."
    STRS(14) = "����ļ���..."
    STRS(15) = "ת����ɾ��ԭ�����ļ�"
    STRS(16) = "ת������ļ��ĸ�ʽ"
    STRS(17) = "ѡ��Ŀ���ĵ��ĸ�ʽ��Ȼ������һ����ʼת����"
    STRS(18) = "ѡ����ת���ĵ���Ŀ���ļ���"
    STRS(19) = "�������ļ���ԭ�ļ�����"
    STRS(20) = "�����������λ��"
    STRS(21) = "����ת���ĵ�"
    STRS(22) = "����ת���ĵ���ʽ��..."
    STRS(23) = "ת�����"
    STRS(24) = "������ת�������ĵ��Ľ��������'״̬'һ����ȷ��ȫ��ת����"
    STRS(25) = "���"
    STRS(26) = "������..."
    
    msgConvertStatus(StatusOK) = "ת����ɣ��� '%1' "
    msgConvertStatus(StatusOKAndDeletefailed) = "�ļ���ת��������ɾ��ԭ�ļ� '%2' ʧ�ܣ��������ļ���ռ�ã�����û��Ȩ�޷��ʸ��ļ���"
    msgConvertStatus(StatusRenamed) = "��ɲ��Ѹ���Ϊ '%1' ��"
    msgConvertStatus(StatusRenamedAndDeletefailed) = "��ɲ��Ѹ���Ϊ '%1' ������ɾ��ԭ�ļ� '%2' ʧ�ܣ��������ļ���ռ�ã�����û��Ȩ�޷��ʸ��ļ���"
    msgConvertStatus(StatusReplaced) = "��ɲ����滻�����ļ� '%1' ��"
    msgConvertStatus(StatusReplacedAndDeletefailed) = "��ɲ����滻�����ļ� '%1' ������ɾ��ԭ�ļ� '%2' ʧ�ܣ��������ļ���ռ�ã�����û��Ȩ�޷��ʸ��ļ���"
    
    msgConvertStatus(StatusFailedforReplace) = "��Ϊ�޷��滻�Ѵ��ڵ��ļ� '%1' ��������"
    msgConvertStatus(StatusFailedforRename) = "��Ϊ���������ļ� '%1' ��������"
    msgConvertStatus(StatusFailedforCancel) = "���û�ȡ��ת����"
    msgConvertStatus(StatusFailedforOpen) = "��ԭ�ļ� '%2' ʱʧ�ܣ��������ļ���ռ�ã������𻵣�����û��Ȩ�޷��ʸ��ļ���"
    msgConvertStatus(StatusFailedforSave) = "ת��ʱ�޷�����Ŀ���ļ� '%2' ������û��Ȩ�޴�ȡ���ļ�����������޷���ȡ��"
    
    STRS(27) = "�����ļ�ת��������£�"
    STRS(28) = "����ת�� %1 ���ļ������� %2 ���ɹ��� %3 ��ʧ�ܡ�"
    STRS(29) = "ת���ɹ���Ϊ %1 %��"
    STRS(29) = "ת���ɹ���Լ %1 %"
    STRS(31) = "�����ļ���Ϊ�����ת��ʧ�ܣ������º˶ԣ�"
    STRS(32) = "û���ļ���Ϊ�����ת��ʧ�ܡ�Nice!"
    STRS(33) = "�ļ�����"
    STRS(34) = "ʧ��ԭ��"
    STRS(35) = "�����ļ�ת��������ϵͳ�Զ�����һЩ�䶯����˶ԣ�"
    STRS(36) = "û���ļ�����Ҫע������"
    STRS(37) = "ϵͳ�ı�"
    STRS(38) = "�����ļ�ת���ɹ���ɣ�"
    STRS(39) = "ת�����"
    STRS(40) = "����û��ת���κ��ļ���"
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
    
    '���ڴ洢��Щ������ʾ��Ҫ����ʾ��
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
