Attribute VB_Name = "mod01"
Option Explicit
Sub test()

Call SjisToUtf8("C:\Users\kyu\Desktop\git_work\helloWorld\mod01.bas", _
    "C:\Users\kyu\Desktop\git_work\helloWorld\mod01utf8.bas")

                
End Sub
Sub SjisToUtf8(a_sFrom, a_sTo)
    Dim streamRead  As New ADODB.Stream '// 読み込みデータ
    Dim streamWrite As New ADODB.Stream '// 書き込みデータ

    '// ファイル読み込み
    streamRead.Type = adTypeText
    streamRead.Charset = "SJIS"
    streamRead.Open
    Call streamRead.LoadFromFile(a_sFrom)
    
    '// ファイル書き込み
    streamWrite.Type = adTypeText
    streamWrite.Charset = "UTF-8"
    streamWrite.Open
    
    '// データ書き込み
    Call streamWrite.WriteText(streamRead.ReadText)
    
    '// 保存
    Call streamWrite.SaveToFile(a_sTo, adSaveCreateOverWrite)
    
    '// クローズ
    streamRead.Close
    streamWrite.Close
End Sub

Public Sub ExportModules()
    '現在のワークブックのモジュールをエクスポートする
    Dim targetModule As VBComponent
    Dim outputPath As String
    Dim fileExt As String
    outputPath = ActiveWorkbook.Path
    For Each targetModule In ActiveWorkbook.VBProject.VBComponents
        fileExt = GetExtFromModuleType(targetModule.Type)
        If fileExt <> "" Then
            ExportModuleWithExt targetModule, outputPath, fileExt
            Debug.Print "Save " & targetModule.Name
        End If
    Next
End Sub

Private Function GetExtFromModuleType(aType As Integer) As String
    '指定されたモジュール・タイプに対応する拡張子を返す
    Select Case aType
    Case vbext_ct_StdModule
        GetExtFromModuleType = "bas"
    Case vbext_ct_ClassModule, vbext_ct_Document
        GetExtFromModuleType = "cls"
    Case vbext_ct_MSForm
        GetExtFromModuleType = "frm"
    End Select
End Function

Private Sub ExportModuleWithExt(aModule As VBComponent, Path As String, Ext As String)
    '指定されたモジュールをエクスポートする
    Dim filePath As String
    Dim saveDir As String
    saveDir = Path & "\" & CreateObject("Scripting.FileSystemObject").GetBaseName(ActiveWorkbook.Name)
    If Dir(saveDir, vbDirectory) = "" Then
        MkDir saveDir
    End If
    filePath = saveDir & "\" & aModule.Name & "." & Ext
    aModule.Export filePath
    Call SjisToUtf8(filePath, filePath)
End Sub
