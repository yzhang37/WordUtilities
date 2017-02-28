Attribute VB_Name = "format_common"
Option Explicit
Const MAX_PATH = 260

Public Const SHGFI_ADDOVERLAYS = &H20
'Version 5.0. Apply the appropriate overlays to the file's icon. The SHGFI_ICON flag must also be set.
Public Const SHGFI_ATTR_SPECIFIED = &H20000
'Modify SHGFI_ATTRIBUTES to indicate that the dwAttributes member of the SHFILEINFO structure at psfi contains the specific attributes that are desired. These attributes are passed to IShellFolder::GetAttributesOf. If this flag is not specified, 0xFFFFFFFF is passed to IShellFolder::GetAttributesOf, requesting all attributes. This flag cannot be specified with the SHGFI_ICON flag.
Public Const SHGFI_ATTRIBUTES = &H800
'Retrieve the item attributes. The attributes are copied to the dwAttributes member of the structure specified in the psfi parameter. These are the same attributes that are obtained from IShellFolder::GetAttributesOf.
Public Const SHGFI_DISPLAYNAME = &H200
'Retrieve the display name for the file, which is the name as it appears in Windows Explorer. The name is copied to the szDisplayName member of the structure specified in psfi. The returned display name uses the long file name, if there is one, rather than the 8.3 form of the file name. Note that the display name can be affected by settings such as whether extensions are shown.
Public Const SHGFI_EXETYPE = &H2000
'Retrieve the type of the executable file if pszPath identifies an executable file. The information is packed into the return value. This flag cannot be specified with any other flags.
Public Const SHGFI_ICON = &H100
'Retrieve the handle to the icon that represents the file and the index of the icon within the system image list. The handle is copied to the hIcon member of the structure specified by psfi, and the index is copied to the iIcon member.
Public Const SHGFI_ICONLOCATION = &H1000
'Retrieve the name of the file that contains the icon representing the file specified by pszPath, as returned by the IExtractIcon::GetIconLocation method of the file's icon handler. Also retrieve the icon index within that file. The name of the file containing the icon is copied to the szDisplayName member of the structure specified by psfi. The icon's index is copied to that structure's iIcon member.
Public Const SHGFI_LARGEICON = &H0
'Modify SHGFI_ICON, causing the function to retrieve the file's large icon. The SHGFI_ICON flag must also be set.
Public Const SHGFI_LINKOVERLAY = &H8000
'Modify SHGFI_ICON, causing the function to add the link overlay to the file's icon. The SHGFI_ICON flag must also be set.
Public Const SHGFI_OPENICON = &H2
'Modify SHGFI_ICON, causing the function to retrieve the file's open icon. Also used to modify SHGFI_SYSICONINDEX, causing the function to return the handle to the system image list that contains the file's small open icon. A container object displays an open icon to indicate that the container is open. The SHGFI_ICON and/or SHGFI_SYSICONINDEX flag must also be set.
Public Const SHGFI_OVERLAYINDEX = &H40
'Version 5.0. Return the index of the overlay icon. The value of the overlay index is returned in the upper eight bits of the iIcon member of the structure specified by psfi. This flag requires that the SHGFI_ICON be set as well.
Public Const SHGFI_PIDL = &H8
'Indicate that pszPath is the address of an ITEMIDLIST structure rather than a path name.
Public Const SHGFI_SELECTED = &H10000
'Modify SHGFI_ICON, causing the function to blend the file's icon with the system highlight color. The SHGFI_ICON flag must also be set.
Public Const SHGFI_SHELLICONSIZE = &H4
'Modify SHGFI_ICON, causing the function to retrieve a Shell-sized icon. If this flag is not specified the function sizes the icon according to the system metric values. The SHGFI_ICON flag must also be set.
Public Const SHGFI_SMALLICON = &H1
'Modify SHGFI_ICON, causing the function to retrieve the file's small icon. Also used to modify SHGFI_SYSICONINDEX, causing the function to return the handle to the system image list that contains small icon images. The SHGFI_ICON and/or SHGFI_SYSICONINDEX flag must also be set.
Public Const SHGFI_SYSICONINDEX = &H4000
'Retrieve the index of a system image list icon. If successful, the index is copied to the iIcon member of psfi. The return value is a handle to the system image list. Only those images whose indices are successfully copied to iIcon are valid. Attempting to access other images in the system image list will result in undefined behavior.
Public Const SHGFI_TYPENAME = &H400
'Retrieve the string that describes the file's type. The string is copied to the szTypeName member of the structure specified in psfi.
Public Const SHGFI_USEFILEATTRIBUTES = &H10
'Indicates that the function should not attempt to access the file specified by pszPath. Rather, it should act as if the file specified by pszPath exists with the file attributes passed in dwFileAttributes. This flag cannot be combined with the SHGFI_ATTRIBUTES, SHGFI_EXETYPE, or SHGFI_PIDL flags.

Public Type SHFILEINFO
    hIcon As Long
    iIcon As Long
    dwAttributes As Long
    szDisplayName As String * MAX_PATH
    szTypeName As String * 80
End Type

Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Public Declare Function SHGetFileInfo Lib "shell32" (ByVal pszPath As String, _
                                                     ByVal wFileAttributes As Long, _
                                                     psfi As SHFILEINFO, _
                                                     ByVal cbFileInfo As Long, _
                                                     ByVal uFlags As Long) As Long

Private Declare Function SetTimer Lib "user32" ( _
    ByVal hWnd As Long, _
    ByVal nIDEvent As Long, _
    ByVal uElapse As Long, _
    ByVal lpTimerFunc As Long) As Long

Private Declare Function KillTimer Lib "user32" ( _
    ByVal hWnd As Long, _
    ByVal nIDEvent As Long) As Long

Dim queryObj As frmQuery
Private thisLeftTime As Long
Dim queryRecommendStrings As String
Dim SameOperationMessage As String
Dim STRS(100) As String

Private Const DEBUG_P = 0
Private Const fc_SKIPDELETEORIGINAL = 1
Private Const fc_SKIPREPLACETARGET = 2
Private Const fc_SKIPSAVEFAILURE = 4
Private Const fc_OPENFAILURE = 8

Private Sub LoadStrings()
    queryRecommendStrings = "����Ҫ׼����ô������"
    SameOperationMessage = "�����Ժ���ļ�������ͬ�Ĳ���"
    STRS(0) = "�洢�ļ����Ѵ����ļ����ֳ�ͻ"
    STRS(1) = "�ļ� '%1' �Ѿ����ڣ���˼�������ת������ֳ�ͻ��"
    STRS(2) = "����Ҫ��������������ֱ���滻�Ѿ����ڵ��ļ���"
    STRS(3) = "������"
    STRS(4) = "�滻"
    STRS(5) = "����"
    STRS(6) = "�滻�ļ����ܾ�"
    STRS(7) = "�滻�Ѿ����ڵ�Ŀ���ļ� '%1' ���ܾ�����������Ϊ����ļ������Ѿ����򿪣������ܵ��˱����Լ�û���㹻��Ȩ���滻����"
    STRS(8) = "����"
    STRS(9) = "�볢�Թرտ���ռ�ø��ļ��ĳ���Ȼ�����ԡ�"
    STRS(10) = "ת���ļ�ʧ��"
    STRS(11) = "��ת����Ŀ���ļ� '%1' ��ʱ�򱻾ܾ����������ڴ洢��λ�ò������㹻��д��Ȩ�ޣ�������𻵡�"
    STRS(12) = "�������Ա����Ȩ�ޣ�����������������һ�Ρ�"
    STRS(13) = "�޷�ɾ��ԭ�ļ�"
    STRS(14) = "ת���ļ���ɣ�����ɾ��ԭ�ļ� '%2' ʧ�ܣ������ļ����ڱ�ռ�á����Ҫɾ���ļ������Եȴ�ת����ɺ���״̬���д���'ע��'�������ļ���"
    STRS(15) = "����ɾ��"
    STRS(16) = "����(�Ƽ�)"
    STRS(17) = "�޷���ԭ�ļ�"
    STRS(18) = "��ԭ�ļ� '%1' ʱ���������⡣�������ļ�������������Ȩ�޲����������Ǵ����𻵵ȶ���ԭ��"
End Sub

Public Function GetPath(ByVal path As String) As String
    Dim i As Long
    i = FindSlash(path)
    If i > 0 Then
        GetPath = Mid(path, 1, FindSlash(path) - 1)
    Else
        GetPath = vbEmpty
    End If
End Function

Public Function GetSuffix(ByVal sFileName As String) As String
    Dim iter As Long
    iter = FindSuffix(sFileName)
    If iter = 0 Then
        GetSuffix = sFileName
    Else
        GetSuffix = Mid(sFileName, iter + 1)
    End If
End Function

Private Function FindSuffix(ByVal sFileName As String) As Long
    Dim iter As Long
    iter = InStrRev(sFileName, ".")
    FindSuffix = iter
End Function

Public Function RemoveSuffix(ByVal sFileName As String) As String
    Dim iter As Long
    iter = FindSuffix(sFileName)
    If iter = 0 Then
        RemoveSuffix = sFileName
    Else
        RemoveSuffix = Left(sFileName, iter - 1)
    End If
End Function

Private Function FindSlash(ByVal path As String) As Long
    Dim iter As Integer
    iter = InStrRev(path, "\")
    If iter = Len(path) Then
        iter = InStrRev(path, "\", iter - 1)
    End If
    FindSlash = iter
End Function

Public Function GetFileName(ByVal path As String) As String
    GetFileName = Mid(path, FindSlash(path) + 1)
End Function

Public Function DocConvert(fsoObject As FileSystemObject, _
                           ByVal FileName As String, _
                           newFileName As String, _
                           FileFormat As Long, _
                           MessageDontShowAgain As Long, _
                           DeleteOriginal As Boolean, _
                           Force As FileForceSaveConstants) As ConvertResultConstants
    
    On Error GoTo Err
    Dim operation As ConvertOperationConstants
    
    'operation ���ڼ�¼��ǰ������ʲô������
    '�����ʱ�����֪������Ϊʲô������ġ�
    operation = ConvertOperationConstants.None
    operation = FileOpen
    
    'Ĭ������£���Ϊû�д���������������˲�����
    Dim renamed As Boolean
    renamed = False
    Dim docu As Document
    
    Dim checked As Boolean, result As Long
OpenOP:
    checked = False
    result = OpenDocument(docu, FileName, MessageDontShowAgain)
    Select Case result
    Case 1
        GoTo OpenOP
    Case 2
        DocConvert = StatusFailedforOpen
        Exit Function
    End Select
    With docu
        'ͨ������£�Force���ޣ���˾���Ҫ�û�ѡ��
        If fsoObject.FileExists(newFileName) Then
            If Force = ConvertOperationConstants.None Then
                checked = False
                result = QueryBox(checked, STRS(0), Replace(STRS(1), "%1", newFileName), STRS(3), STRS(4), STRS(5), 3, True, 7)
                Select Case result
                Case 1  'Rename
                    newFileName = AdaptFileName(fsoObject, newFileName)
                    renamed = True
                    If (checked) Then
                        Force = ForceRename
                    End If
                    'Ԥ���ý��״̬��
                    DocConvert = StatusRenamed
                Case 2  'Replace
                    If (checked) Then
                        Force = ForceReplace
                    End If
                    GoTo ReplaceOP
                Case 3  'Skip
                    If (checked) Then
                        Force = ForceSkip
                    End If
                    'Ԥ���ý��״̬��: �滻Ŀ���ļ�
                    GoTo SKipOP
                End Select
            ElseIf Force = ForceRename Then
RenameOP:
                '�Զ�������һ�����Դ�ŵ��ļ���
                newFileName = AdaptFileName(fsoObject, newFileName)
                
                'Ԥ���ý��״̬��
                DocConvert = StatusRenamed
                
                renamed = True
            ElseIf Force = ForceReplace Then
ReplaceOP:
                'Ԥ���ý��״̬��: �滻Ŀ���ļ�
                DocConvert = StatusReplaced
                
                operation = ConvertOperationConstants.ReplaceTargetFile
                
                result = ReplaceExistFile(newFileName, MessageDontShowAgain)
                Select Case result
                Case 1
                    GoTo ReplaceOP
                Case 2
                    GoTo RenameOP
                Case 3
                    DocConvert = ConvertResultConstants.StatusFailedforReplace
                    GoTo AfterSave
                End Select
            Else
SKipOP:
                'Ԥ���ý��״̬��: �滻Ŀ���ļ�
                DocConvert = StatusFailedforRename
            
                '��������ļ�
                GoTo AfterSave
            End If
        Else
            'Ԥ���ý��״̬��
            DocConvert = StatusOK
        End If
SaveOp:
        result = SaveAsNewFile(docu, newFileName, FileFormat, MessageDontShowAgain)
        Select Case result
        Case 1
            GoTo SaveOp
        Case 2
            DocConvert = StatusFailedforSave
            GoTo AfterSave
        End Select
        'After Save Operation
AfterSave:
        .Close
    End With
    
    
    '��� DeleteOriginal ���棬��ɾ��ԭ�ļ�����
    If DeleteOriginal = True And DocConvert <= StatusNormal Then
DeleteOrigin:
        result = RemoveOriginalFile(FileName, MessageDontShowAgain)
        Select Case result
        Case 1
            GoTo DeleteOrigin
        Case 2
            DocConvert = DocConvert + 1
        End Select
    End If
    
    Exit Function
Err:
    Select Case Err.Number
    Case 75, 70 'Access to file failed
        Dim err_checked As Boolean
        Select Case operation
        Case ConvertOperationConstants.FileOpen
            'TODO: ���ļ���ʱ��ͳ���������
        End Select
    Case Else '//default:
        MsgBox "��������" + Str(Err.Number) + vbLf + Err.Description
    End Select
End Function

Private Function OpenDocument(OpenDoc As Document, FileName As String, MessageDontShowAgain As Long) As Long
    On Error GoTo Err
    Set OpenDoc = Documents.Open(FileName, False, True, False, _
                  Visible:=False)
    OpenDocument = 0
    Exit Function
Err:
    Dim err_checked As Boolean, result As Long
    result = 2
    If (MessageDontShowAgain And fc_OPENFAILURE) = 0 Then
        Dim errorMessage As String
        errorMessage = Replace(STRS(18), "%1", FileName)
        If Err.Number <> 70 And Err.Number <> 75 Then
            errorMessage = errorMessage & ExtractError(Err)
        End If
        
        result = QueryBox(err_checked, STRS(17), errorMessage, _
                          STRS(8), STRS(5), , 2, True, 2, STRS(12))
        If err_checked Then MessageDontShowAgain = MessageDontShowAgain Or fc_OPENFAILURE
    End If
    OpenDocument = result
End Function

Private Function RemoveOriginalFile(oldFileName As String, MessageDontShowAgain As Long) As Long
    On Error GoTo Err:
    Kill oldFileName
    RemoveOriginalFile = 0
    Exit Function
Err:
    Dim err_checked As Boolean, result As Long
    result = 3
    If (MessageDontShowAgain And fc_SKIPDELETEORIGINAL) = 0 Then
        Dim errorMessage As String
        errorMessage = Replace(STRS(14), "%2", oldFileName)
        
        If Err.Number <> 70 And Err.Number <> 75 Then
            errorMessage = errorMessage & ExtractError(Err)
        End If
        result = QueryBox(err_checked, STRS(13), errorMessage, _
                          STRS(15), STRS(16), , 2, True, 2, STRS(9))
        If err_checked Then MessageDontShowAgain = MessageDontShowAgain Or fc_SKIPDELETEORIGINAL
    End If
    RemoveOriginalFile = result
End Function

Private Function SaveAsNewFile(doc As Document, newFileName As String, FileFormat As Long, MessageDontShowAgain As Long) As Long
    On Error GoTo Err:
    doc.SaveAs FileName:=newFileName, FileFormat:=FileFormat
    SaveAsNewFile = 0
    Exit Function
Err:
    Dim err_checked As Boolean, result As Long
    result = 2
    If (MessageDontShowAgain And fc_SKIPSAVEFAILURE) = 0 Then
        Dim errorMessage As String
        errorMessage = Replace(STRS(11), "%1", newFileName)
        If Err.Number <> 70 And Err.Number <> 75 Then
            errorMessage = ExtractError(Err)
        End If
        
        result = QueryBox(err_checked, STRS(10), errorMessage, _
                          STRS(8), STRS(5), , 2, True, 2, STRS(12))
        If err_checked Then MessageDontShowAgain = MessageDontShowAgain Or fc_SKIPSAVEFAILURE
    End If
    SaveAsNewFile = result
End Function

Private Function ReplaceExistFile(newFileName As String, MessageDontShowAgain As Long) As Long
    On Error GoTo Err:
    Kill newFileName
    ReplaceExistFile = 0
    Exit Function
Err:
    Dim err_checked As Boolean, result As Long
    result = 3
    If (MessageDontShowAgain And fc_SKIPREPLACETARGET) = 0 Then
        Dim errorMessage As String
        errorMessage = Replace(STRS(7), "%1", newFileName)
        
        If Err.Number <> 70 And Err.Number <> 75 Then
            errorMessage = errorMessage & ExtractError(Err)
        End If
        result = QueryBox(err_checked, STRS(6), errorMessage, _
                          STRS(8), STRS(3), STRS(5), 3, True, 4, STRS(9))
        If err_checked Then MessageDontShowAgain = MessageDontShowAgain Or fc_SKIPREPLACETARGET
    End If
    ReplaceExistFile = result
End Function

Public Function QueryBox(checkedStatus As Boolean, queryTitle As String, queryMessage As String, _
                         Optional b1msg As String = "", Optional b2msg As String = "", Optional b3msg As String = "", _
                         Optional ButtonCount As Integer = 3, Optional DisplayRepeatOption As Boolean = True, _
                         Optional Filter As Long = &H7, Optional queryRecommend As String = "", _
                         Optional Same As String = "") As Long
    Dim fQueryInst As New frmQuery
    Load fQueryInst
    With fQueryInst
        .Caption = queryTitle
        .msg1.Caption = queryMessage
        If Len(queryRecommend) > 0 Then .msg2.Caption = queryRecommend
        If Len(Same) > 0 Then .cbxSame.Caption = Same
        .btn1.Caption = b1msg
        .btn2.Caption = b2msg
        .btn3.Caption = b3msg
        .ButtonCount = ButtonCount
        .RepeatOptionVisible = DisplayRepeatOption
        .OptionFilter = Filter
        .Show vbModal
        checkedStatus = .IsSameClicked
        QueryBox = .value
    End With
    Unload fQueryInst
    Set fQueryInst = Nothing
End Function

Private Function AdaptFileName(fsoObject As FileSystemObject, FileName As String) As String
    Dim newFileName As String
    newFileName = FileName
    Dim i As Long, j As Long
    j = InStrRev(FileName, "\")
    j = InStr(j + 1, FileName, ".") - 1
    If j <= 0 Then j = Len(FileName)
    i = 2
    While fsoObject.FileExists(newFileName)
        newFileName = Mid(FileName, 1, j) + " (" + Trim(Str(i)) + ")" + Mid(FileName, j + 1)
        i = i + 1
    Wend
    AdaptFileName = newFileName
End Function

Public Sub StartConvUI()
    LoadStrings
    frmConvOption.Show
End Sub

Private Function ExtractError(exError As ErrObject) As String
    ExtractError = IIf(DEBUG_P = 1, Trim(Str(exError.Number)) + ": ", vbNullString) + exError.Description
End Function

