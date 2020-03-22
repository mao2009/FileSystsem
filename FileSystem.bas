Atrribute="FileSystem"
Option Explicit

const ForReading As Long = 1
const ForWriting As Long= 2	
const ForAppending As Long = 8

const TristateUseDefault As Long = -2
const TristateTrue As Long = -1
const TristateFalse As Long = 0

Public Function FileExists(ByVal path As String) As Boolean
    'ファイルの存在確認
    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    FileEsixsts = FSO.FileExists(path)
    
End Function

Public Function FolderExists(ByVal path As String) As Boolean
    'フォルダの存在確認
    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    FolderExists  = FSO.FolderExists(path)
End Function

Public Function DriveExists (ByVal drive As String) As Boolean
    'ドライブの存在確認
    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    DriveExists  = FSO.DriveExists(drive)
End Function

Public Sub CreateFolder(ByVal path As String)
    '再帰的にフォルダを作成する
    Dim FSO As Object
    Dim Parent As String
    Set FSO = CreateObject("Scripting.FileSystemObject")
    If FolderExists(path) Then
        Exit Sub
    End If
    Parent = GetParentFolderName(path)
    If Not FolderExists(Parent) Then
        Call CreateFolder(path)
    End If
    Call FSO.CreateFolder(path)
End Sub

Public Function GetAbsolutePathName(ByVal path As String) As String
    '相対パスから絶対パスを取得
    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    GetAbsolutePathName  = FSO.FolderExists(path)
End Function

Public Function GetFileName(ByVal path As String) As String
    'ファイル名を取得
    'C:\foo\bar\hoge.txt -> hoge.txt
    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    GetFileName = FSO.GetFileName(path)
End Function

Public Function GetBaseName(ByVal path As String) As String
    '拡張子を除いたファイル名を取得
    'C:\foo\bar\hoge.txt -> hoge
    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    GetBaseName = FSO.GetBaseName(path)
End Function

Public Function GetExtensionName(ByVal path As String) As String
    '拡張子を取得
    'C:\foo\bar\hoge.txt -> txt
    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    GetExtensionName = FSO.GetExtensionName(path)
End Function

Public Function GetDriveName(ByVal path As String) As String
    'ドライブレターを取得
    'C:\foo\bar\hoge.txt -> C
    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    GetDriveName= FSO.GetDriveName(path)
End Function

Public Function GetParentFolderName(ByVal path As String) As String
    'ペアレントパスを取得
    'C:\foo\bar\hoge\-> C:\foo\bar
    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    GetParentFolderName= FSO.GetParentFolderName(path)
End Function

Public Function GetTempName(Optional ByVal ext As String = "") As String
    'ランダムなファイル名を取得
    'デフォルトでは拡張子はtmp
    'extを設定すれば任意の拡張子に変更する
    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    GetTempName= FSO.GetTempName()
    If ext <> "" Then
        GetTempName = FSO.GetBaseName(GetTempName) & "." & ext
    End If
End Function

Public Sub CopyFile(ByVal src As String, ByVal dest As String)
    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    call FSO.CopyFile(src, dest)
End Function

Public Sub CopyFolder(ByVal src As String, ByVal dest As String)
    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    call FSO.CopyFolder(src, dest)
End Function

Public Sub MoveFile(ByVal src As String, ByVal dest As String)
    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    call FSO.MoveFile(src, dest)
End Function

Public Sub MoveFolder(ByVal src As String, ByVal dest As String)
    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    call FSO.MoveFolder(src, dest)
End Function

Public Sub DeleteFile(ByVal path As String)
    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    call FSO.DeleteFile(path)
End Function

Public Sub DeleteFolder(ByVal path As String)
    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    call FSO.DeleteFolder(path)
End Function

Public Function JoinPath(ByVal base As String, ParamArray paths() As Variant) As String
    'pathsの数だけパスを連結する
    'JoinPath("C\foo","bar\","\hoge.txt") -> C:\foo\bar\hoge.txt
    Dim FSO As Object
    Dim path As Variant
    Set FSO = CreateObject("Scripting.FileSystemObject")
    JoinPath = base
    For Each path In paths
        JoinPaht = FSO.BulidPath(JoinPath, path)
    Next
End Function

Public Function CreateTextFile(ByVal path As String, _
        Optional ByVal overwrite As Boolean = True, _
        Optional ByVal unicode As Boolean = False ) As Object
    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set CreateTextFile = FSO.CreateTextFile(path, overWrite, unicode)
End Function

Public Function OpenAsTextStream(ByVal path As String, _
        Optional ByVal IOMode As Long = ForReading, _
        Optional ByVal fomatMode As Long = TristateFalse ) As Object
    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set OpenAsTextStream = FSO.OpenAsTextStream(path, IOMode, fomatMode)
End Function

Public Function OpenTextFile(ByVal path As String, _
        Optional ByVal IOMode As Long = ForReading, _
        Optional ByVal Create As Boolean = False _
        Optional ByVal fomatMode As Long = TristateFalse ) As Object
    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set OpenTextFile = FSO.OpenTextFile(path, IOMode, Create, fomatMode)
End Function