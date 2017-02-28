VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FileSave"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Enum FileForceSaveConstants
    None = 0
    ForceSkip = 1
    ForceRename = 2
    ForceReplace = 3
End Enum

Public Enum ConvertOperationConstants
    None = 0
    FileOpen = 1
    ReplaceTargetFile = 2
    FileSave = 3
    DeleteOriginalFile = 4
End Enum

Public Enum ConvertResultConstants
    StatusOK = 0
    StatusOKAndDeletefailed = 1
    StatusRenamed = 2
    StatusRenamedAndDeletefailed = 3
    StatusReplaced = 4
    StatusReplacedAndDeletefailed = 5
    StatusPerfect = 0
    StatusCaution = 1
    StatusNormal = StatusReplacedAndDeletefailed

    StatusFailedforCancel = 6
    StatusFailedforRename = 7
    StatusFailedforOpen = 8
    StatusFailedforSave = 9
    StatusFailedforReplace = 10
    StatusFailed = StatusFailedforCancel
    
End Enum
