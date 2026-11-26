Attribute VB_Name = "MTChapSyncMod"
Option Explicit

'=============================================================
' MathType 公式章号自动同步 (MathType Chapter Number Sync)
'
' 原理:
'   MathType 的公式编号是一个域结构:
'     { MACROBUTTON MTPlaceRef
'         { SEQ MTEqn \h }
'         ( { SEQ MTChap \c } . { SEQ MTEqn \c } ) }
'   其中计 "章号" 的是 SEQ MTChap 这个 Word 序列表域。
'   只要在每个章标题后面放一个 { SEQ MTChap \r N \h } 的
'   隐藏锚点, \c 就会显示当前章号, 而 MathType 的
'   插入编号 / 更新公式编号 / 交叉引用 全部照常工作。
'
' 本模块只做两件事:
'   1) 维护锚点 (每个章标题后恰好一个, 幂等)
'   2) 刷新 SEQ MTChap \c 的显示值
' 不改动 MathType 的编号域, 因此编号样式完全跟随
' MathType "格式化公式编号" 里的设置。
'=============================================================

Private Const APPKEY As String = "MTChapSync"
Private Const SECCFG As String = "Cfg"

Private m_Busy As Boolean
Private m_LastTick As Double
Public g_LastChanged As Boolean

'-------------------------------------------------------------
' Word 命令宏: 接管 文件->保存 / 打印, 保存前自动同步
' 全局模板里定义 FileSave 后, 用户按 Ctrl+S 或点保存都会先走这里
'-------------------------------------------------------------
Public Sub FileSave()
    Dim d As Document
    Dim n As Long
    Dim msg As String
    On Error Resume Next
    Set d = ActiveDocument
    On Error GoTo 0
    If d Is Nothing Then Exit Sub

    If MTChapSync_IsOn("AutoAll") Then
        Application.StatusBar = "正在同步公式编号..."
        On Error Resume Next
        n = MTChapSync_Sync(d, True)
        On Error GoTo 0

        If n > 0 Then
            msg = "公式编号已同步：" & n & " 章" & IIf(g_LastChanged, "（有更新）", "") & "  " & Format$(Now, "hh:mm:ss")
        Else
            msg = ""
        End If
    End If

    d.Save          ' 让 Word 自己处理"未命名/只读"等常规错误

    ' 保存完成后再写状态栏，否则会被 Word 自己的"正在保存"覆盖
    If Len(msg) > 0 Then Application.StatusBar = msg
End Sub

Public Sub FileSaveAs()
    Dim d As Document
    Dim n As Long
    On Error Resume Next
    Set d = ActiveDocument
    On Error GoTo 0
    If d Is Nothing Then Exit Sub
    If MTChapSync_IsOn("AutoAll") Then
        On Error Resume Next
        n = MTChapSync_Sync(d, True)
        On Error GoTo 0
    End If
    Application.Dialogs(wdDialogFileSaveAs).Show
End Sub

Public Sub FilePrint()
    Dim d As Document
    Dim n As Long
    On Error Resume Next
    Set d = ActiveDocument
    On Error GoTo 0
    If d Is Nothing Then Exit Sub
    If MTChapSync_IsOn("AutoAll") Then
        On Error Resume Next
        n = MTChapSync_Sync(d, True)
        On Error GoTo 0
    End If
    Application.PrintOut
End Sub

Public Sub AutoExec()
    On Error Resume Next
    SaveSetting APPKEY, SECCFG, "AutoExecAt", CStr(Now)
End Sub

'-------------------------------------------------------------
' 功能区按钮回调
'-------------------------------------------------------------
Public Sub MTChapSync_OnRibbonSync(ByVal control As IRibbonControl)
    On Error Resume Next
    MTChapSync_Manual
End Sub

'-------------------------------------------------------------
' 自动同步总开关(保存+打印)
'-------------------------------------------------------------
Public Function MTChapSync_IsOn(ByVal key As String) As Boolean
    On Error Resume Next
    MTChapSync_IsOn = (GetSetting(APPKEY, SECCFG, key, "1") = "1")
End Function

Public Sub MTChapSync_Flip(ByVal key As String)
    Dim v As String
    On Error Resume Next
    v = GetSetting(APPKEY, SECCFG, key, "1")
    SaveSetting APPKEY, SECCFG, key, IIf(v = "1", "0", "1")
End Sub

Public Sub MTChapSync_AutoOn(ByVal control As IRibbonControl)
    On Error Resume Next
    SaveSetting APPKEY, SECCFG, "AutoAll", "1"
    Application.StatusBar = "公式编号：保存/打印时自动同步 = 开"
End Sub

Public Sub MTChapSync_AutoOff(ByVal control As IRibbonControl)
    On Error Resume Next
    SaveSetting APPKEY, SECCFG, "AutoAll", "0"
    Application.StatusBar = "公式编号：保存/打印时自动同步 = 关"
End Sub

'-------------------------------------------------------------
' 手动触发(带提示)
'-------------------------------------------------------------
Public Sub MTChapSync_Manual()
    Dim n As Long
    On Error Resume Next

    If Application.Documents.Count = 0 Then Exit Sub
    If ActiveDocument.ReadOnly Then
        Application.StatusBar = "公式编号：当前文档只读，无法写入锚点"
        Exit Sub
    End If

    n = MTChapSync_Sync(ActiveDocument, True)

    Select Case n
        Case Is > 0
            Application.StatusBar = "公式编号已同步：" & n & " 章" & IIf(g_LastChanged, "（有更新）", "") & "  " & Format$(Now, "hh:mm:ss")
        Case -1
            Application.StatusBar = "公式编号：本文档没有 MathType 编号"
        Case -2
            Application.StatusBar = "公式编号：没有找到带自动编号的一级标题"
        Case Else
            Application.StatusBar = "公式编号：未执行（正忙）"
    End Select
End Sub

'-------------------------------------------------------------
' 无界面入口(给自动化/测试用)
'-------------------------------------------------------------
Public Function MTChapSync_Run() As Long
    On Error Resume Next
    MTChapSync_Run = MTChapSync_Sync(ActiveDocument, True)
End Function

'-------------------------------------------------------------
' 核心同步
'   返回 >=0 同步的章节数
'        -1 文档没有 MathType 编号
'        -2 没有找到章标题
'        -3 正忙/参数问题
'-------------------------------------------------------------
Public Function MTChapSync_Sync(ByVal doc As Document, ByVal silent As Boolean) As Long
    Dim idx() As Long
    Dim num() As Long
    Dim cnt As Long
    Dim p As Paragraph
    Dim f As Field
    Dim i As Long
    Dim needSec As Boolean
    Dim hasMT As Boolean
    Dim anchors As String
    Dim desired As String
    Dim valsOk As Boolean
    Dim cur As Long
    Dim code As String

    On Error GoTo Failed

    MTChapSync_Sync = -3
    g_LastChanged = False
    If doc Is Nothing Then Exit Function
    If m_Busy Then Exit Function
    If doc.ReadOnly Then Exit Function

    m_Busy = True
    MTChapSync_Sync = -1

    ' --- 1) 章标题: 先看大纲级别(便宜), 只有 1 级标题才去取编号(昂贵) ---
    For Each p In doc.Paragraphs
        If p.Range.ParagraphFormat.OutlineLevel = wdOutlineLevel1 Then
            If Len(Trim$(p.Range.ListFormat.ListString)) > 0 Then
                cnt = cnt + 1
                ReDim Preserve idx(1 To cnt)
                ReDim Preserve num(1 To cnt)
                idx(cnt) = p.Range.End
                num(cnt) = ChapterNumber(p, cnt)
                desired = desired & CStr(num(cnt)) & ","
            End If
        End If
    Next p

    If cnt = 0 Then
        MTChapSync_Sync = -2
        GoTo Done
    End If

    ' --- 2) 一次扫描域: 锚点序列 / 显示值是否落后 / 有无 MT / 是否用 MTSec ---
    cur = 0
    valsOk = True
    For Each f In doc.Fields
        code = f.Code.Text
        If InStr(1, code, "MTPlaceRef", vbTextCompare) > 0 Then hasMT = True
        If InStr(1, code, "MTSec", vbTextCompare) > 0 Then needSec = True
        If InStr(1, code, "MTChap", vbTextCompare) > 0 Then
            If InStr(1, code, "MACROBUTTON", vbTextCompare) = 0 Then
                If InStr(1, code, "\r", vbTextCompare) > 0 Then
                    cur = AnchorNumber(code)
                    anchors = anchors & CStr(cur) & ","
                ElseIf InStr(1, code, "\c", vbTextCompare) > 0 Then
                    If f.Result.Text <> CStr(cur) Then valsOk = False
                End If
            End If
        End If
    Next f

    If Not hasMT Then GoTo Done          ' 返回 -1

    ' --- 3) 快路径: 结构和显示值都没变, 一个字节都不改 ---
    If anchors = desired Then
        If valsOk Then
            MTChapSync_Sync = cnt
            GoTo Done
        End If
        UpdateStaleChapterFields doc, True
        g_LastChanged = True
        MTChapSync_Sync = cnt
        GoTo Done
    End If

    ' --- 4) 结构变了: 重建锚点 ---
    RemoveAnchors doc, "MTChap"
    If needSec Then RemoveAnchors doc, "MTSec"

    For i = cnt To 1 Step -1
        AddAnchor doc, idx(i), "MTChap", num(i)
    Next i

    UpdateStaleChapterFields doc, False
    g_LastChanged = True

    MTChapSync_Sync = cnt

Done:
    m_Busy = False
    Exit Function

Failed:
    m_Busy = False
    If Not silent Then
        MsgBox "同步失败: " & Err.Description, vbExclamation, "公式编号同步"
    End If
End Function

'-------------------------------------------------------------
' 内部: 从 SEQ 域代码里取出 \r 后面的数字
'-------------------------------------------------------------
Private Function AnchorNumber(ByVal code As String) As Long
    Dim i As Long
    Dim s As String
    Dim ch As String

    i = InStr(1, code, "\r", vbTextCompare)
    If i = 0 Then Exit Function
    For i = i + 2 To Len(code)
        ch = Mid$(code, i, 1)
        If ch >= "0" And ch <= "9" Then
            s = s & ch
        ElseIf Len(s) > 0 Then
            Exit For
        End If
    Next i
    If Len(s) > 0 And Len(s) <= 4 Then AnchorNumber = CLng(s)
End Function

'-------------------------------------------------------------
' 内部: 刷新显示值; onlyStale=True 时只更新已经不对的
'-------------------------------------------------------------
Private Sub UpdateStaleChapterFields(ByVal doc As Document, ByVal onlyStale As Boolean)
    Dim f As Field
    Dim code As String
    Dim cur As Long

    On Error Resume Next
    For Each f In doc.Fields
        code = f.Code.Text
        If InStr(1, code, "MTChap", vbTextCompare) > 0 Then
            If InStr(1, code, "MACROBUTTON", vbTextCompare) = 0 Then
                If InStr(1, code, "\r", vbTextCompare) > 0 Then
                    cur = AnchorNumber(code)
                ElseIf InStr(1, code, "\c", vbTextCompare) > 0 Then
                    If (Not onlyStale) Or (f.Result.Text <> CStr(cur)) Then f.Update
                End If
            End If
        End If
    Next f
End Sub

'-------------------------------------------------------------
' 光标触发: 只在"当前段是章标题"或"上一段是章标题"时执行
'-------------------------------------------------------------
Public Sub MTChapSync_CheckSelection(ByVal Sel As Selection)
    Dim p As Paragraph
    Dim prev As Paragraph

    On Error Resume Next
    If m_Busy Then Exit Sub
    If Sel Is Nothing Then Exit Sub
    If Sel.Range Is Nothing Then Exit Sub
    If Timer - m_LastTick < 2# Then Exit Sub

    Set p = Sel.Paragraphs(1)
    If p Is Nothing Then Exit Sub
    If p.Range.ListFormat Is Nothing Then Exit Sub

    If IsChapterParagraph(p) Then
        m_LastTick = Timer
        MTChapSync_Sync Sel.Document, True
        Exit Sub
    End If

    Set prev = p.Previous
    If prev Is Nothing Then Exit Sub
    If IsChapterParagraph(prev) Then
        m_LastTick = Timer
        MTChapSync_Sync Sel.Document, True
    End If
End Sub

'-------------------------------------------------------------
' 内部: 判断段落是否为"章"
'-------------------------------------------------------------
Private Function IsChapterParagraph(ByVal p As Paragraph) As Boolean
    On Error Resume Next
    IsChapterParagraph = False
    If p.Range.ListFormat Is Nothing Then Exit Function
    If p.Range.ParagraphFormat.OutlineLevel <> wdOutlineLevel1 Then Exit Function
    If Len(Trim$(p.Range.ListFormat.ListString)) = 0 Then Exit Function
    IsChapterParagraph = True
End Function

'-------------------------------------------------------------
' 内部: 取章号 (优先用标题自身的编号, 取不到则用序号)
'-------------------------------------------------------------
Private Function ChapterNumber(ByVal p As Paragraph, ByVal ordinal As Long) As Long
    Dim s As String
    Dim part As String
    Dim i As Long
    Dim digits As String

    On Error Resume Next
    ChapterNumber = ordinal

    s = Trim$(p.Range.ListFormat.ListString)
    If Len(s) = 0 Then Exit Function

    part = Split(s, ".")(0)
    For i = 1 To Len(part)
        If Mid$(part, i, 1) >= "0" And Mid$(part, i, 1) <= "9" Then
            digits = digits & Mid$(part, i, 1)
        End If
    Next i
    If Len(digits) > 0 And Len(digits) <= 3 Then
        ChapterNumber = CLng(digits)
    End If
End Function

'-------------------------------------------------------------
' 内部: 插入一个 SEQ 锚点
'-------------------------------------------------------------
Private Sub AddAnchor(ByVal doc As Document, ByVal pos As Long, ByVal seqName As String, ByVal n As Long)
    Dim r As Range
    Dim f As Field

    On Error Resume Next
    If pos <= 0 Then Exit Sub
    If pos >= doc.Content.End - 1 Then Exit Sub

    Set r = doc.Range(pos, pos)
    Set f = doc.Fields.Add(r, wdFieldEmpty, " SEQ " & seqName & " \r " & n & " \h ", False)
    If Not f Is Nothing Then f.Update
End Sub

'-------------------------------------------------------------
' 内部: 删除某一序列的所有锚点(保留 \c 显示域)
'-------------------------------------------------------------
Private Sub RemoveAnchors(ByVal doc As Document, ByVal seqName As String)
    Dim i As Long
    Dim f As Field
    Dim code As String

    On Error Resume Next
    For i = doc.Fields.Count To 1 Step -1
        Set f = doc.Fields(i)
        If Not f Is Nothing Then
            code = f.Code.Text
            If InStr(1, code, seqName, vbTextCompare) > 0 Then
                If InStr(1, code, "\c", vbTextCompare) = 0 Then
                    If InStr(1, code, "\r", vbTextCompare) > 0 _
                       Or InStr(1, code, "\h", vbTextCompare) > 0 _
                       Or InStr(1, code, "\s", vbTextCompare) > 0 Then
                        f.Delete
                    End If
                End If
            End If
        End If
    Next i
End Sub

