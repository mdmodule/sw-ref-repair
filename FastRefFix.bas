' ******************************************************************************
' SolidWorks 2025 极速引用修复宏 (不加载图形版)
' 逻辑：模仿"参考"对话框逻辑，直接修改文件头索引。
' 优点：极速、不吃内存、支持子文件夹、结束后有统计提示。
' ******************************************************************************

Option Explicit

Dim swApp As SldWorks.SldWorks
Dim fso As Object
Dim totalDrw As Long ' 扫描到的工程图总数
Dim fixedLinks As Long ' 成功修复的链接总数

Sub main()
 Dim ShellApp As Object
 Dim Folder As Object
 Dim folderPath As String
 
 Set swApp = Application.SldWorks
 Set fso = CreateObject("Scripting.FileSystemObject")
 
 ' 初始化统计
 totalDrw = 0
 fixedLinks = 0
 
 ' 1. 选择根目录
 Set ShellApp = CreateObject("Shell.Application")
 Set Folder = ShellApp.BrowseForFolder(0, "选择包含工程图和模型的根目录", 0)
 
 If Folder Is Nothing Then Exit Sub
 folderPath = Folder.Items.Item.path & "\"
 
 ' 2. 开始极速修复逻辑
 RecursiveFix folderPath
 
 ' 3. 运行结束弹出提示框
 MsgBox "【极速修复完成】" & vbCrLf & vbCrLf & _
 "扫描工程图总数: " & totalDrw & " 个" & vbCrLf & _
 "成功修复链接数: " & fixedLinks & " 处" & vbCrLf & vbCrLf & _
 "提示：若修复数为0，请检查工程图与模型文件名是否完全一致。", _
 vbInformation, "SolidWorks 自动化报告"
End Sub

' --- 递归函数：穿透所有子文件夹 ---
Sub RecursiveFix(ByVal path As String)
 Dim parentFolder As Object, subFolder As Object, file As Object
 Set parentFolder = fso.GetFolder(path)
 
 ' 处理当前文件夹的文件
 For Each file In parentFolder.Files
 ' 只针对工程图进行索引检查
 If LCase(fso.GetExtensionName(file.Name)) = "slddrw" Then
 totalDrw = totalDrw + 1
 ProcessReferences file.path, path
 End If
 Next file
 
 ' 进入每一个子文件夹
 For Each subFolder In parentFolder.SubFolders
 RecursiveFix subFolder.path & "\"
 Next subFolder
End Sub

' --- 核心逻辑：文件头重定向 (不打开文件) ---
Sub ProcessReferences(drwPath As String, currentDir As String)
 Dim vRefs As Variant
 Dim i As Long
 Dim oldPath As String, drwBaseName As String
 Dim targetPath As String
 Dim status As Boolean
 
 ' 获取工程图不带后缀的名字，例如 "脚轮 5013..."
 drwBaseName = fso.GetBaseName(drwPath)
 
 ' 获取该工程图依赖的所有模型路径 (关键 API)
 vRefs = swApp.GetDocumentDependencies2(drwPath, True, True, False)
 
 If IsEmpty(vRefs) Then Exit Sub
 
 ' vRefs 数组中，奇数位是旧的引用路径
 For i = 1 To UBound(vRefs) Step 2
 oldPath = vRefs(i)
 
 ' 检查：如果旧模型名字和现在工程图名字对不上
 If LCase(fso.GetBaseName(oldPath)) <> LCase(drwBaseName) Then
 
 ' 在当前文件夹寻找同名模型
 targetPath = ""
 If fso.FileExists(currentDir & drwBaseName & ".SLDPRT") Then
 targetPath = currentDir & drwBaseName & ".SLDPRT"
 ElseIf fso.FileExists(currentDir & drwBaseName & ".SLDASM") Then
 targetPath = currentDir & drwBaseName & ".SLDASM"
 End If
 
 ' 如果找到了匹配的新模型，执行"重定向"
 If targetPath <> "" And LCase(targetPath) <> LCase(oldPath) Then
 ' 直接改写文件索引，不加载模型
 status = swApp.ReplaceReferencedDocument(drwPath, oldPath, targetPath)
 If status Then fixedLinks = fixedLinks + 1
 End If
 End If
 Next i
End Sub
