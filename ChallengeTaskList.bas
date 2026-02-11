Attribute VB_Name = "ChallengeTaskList"
'《ChallengeTask任务一览》
'accuracyBelowAverage: 练习准确率低于题库平均正确率的题目
'classicShuffle:传统模板随机练习功能
'recentFault:练习最近做错的题目，默认30天内的错题，可视情况调整
'unPracticed7days:练习7天内未做过的题目
'valueBelowAverage：练习价值系数高于题库平均价值系数的题目
'highValue:练习高价值系数（高难）题目
'statMasteryUpByPercentage '多次练习使掌握度提升（低于50为2%，低于60为1%，低于70为0.5%，低于80为0.3%，高于80为0.15%）
'statLearnTime20min '当日练习时长达到20分钟
'statFaultAccuracy '练习2次任何时刻的错题，但要求总体正确率达到80%
'perfectChallenge '完美挑战，从valueBelowAverage函数里选出15题，共5个答错次数，单题正确率<70%判答错。答错后仍旧可以继续答题，但本次挑战失败


Public LogContent As String
Public taskAccomplished As Boolean
Public logline() As String
Public lineBlock() As String
Public CheckTaskNum() As String
Public universalTaskNum As Integer
Public maxTopicNum As Integer
Public isFromChallengeTask As Boolean
Public currentPracticedNum As Integer

Public Function removeChallengelib(libname As String)
    If Dir(App.Path & "\Archive\" & libname & "\ChallengeLib") <> "" Then
        Kill (App.Path & "\Archive\" & libname & "\ChallengeLib")
        Kill (App.Path & "\Archive\" & libname & "\ChallengeLog")
        Kill (App.Path & "\Archive\" & libname & "\ChallengeTasks")
    End If
    ChallengeLib = ""
    If Dir(App.Path & "\ChallengeLib") <> "" Then Kill App.Path & "\ChallengeLib"
End Function


Public Function punchStatCheck() As Boolean
If Dir(App.Path & "\punchToday") = "" Then
    punchStatCheck = False
    Exit Function
End If
OpenTxt App.Path & "\punchToday"
InputStr = DecryptString(InputStr, UpdateKey)
InputStr = Replace$(InputStr, vbCrLf, "")
Dim block() As String, unit() As String
block = Split(InputStr, "\\")
unit = Split(block(UBound(block) - 1), "::")
Dim verif As Date
verif = Now
If unit(0) = verif Then
    punchStatCheck = True
Else
    punchStatCheck = False
End If
End Function

Public Function punchToday(Optional ByRef rtnTime As Integer, Optional ByRef rtnCount As Integer) '每次考试练习结束后的“保存进度”性质的行为
Dim a As Object, fs As Object
Set fs = CreateObject("Scripting.FileSystemObject")
Dim fso As New FileSystemObject
Dim folderName As String
If Dir(App.Path & "\punchToday") = "" Then
    Set a = fs.CreateTextFile(App.Path & "\punchToday", True)
    a.WriteLine "" '默认练习强度
    a.Close
    Set a = Nothing
End If
Dim getCount As Integer, getTime As Integer, totalCount As Integer, totalTime As Integer
OpenTxt App.Path & "\punchToday"
totalCount = 0
totalTime = 0
If Len(Replace(InputStr, vbCrLf, "")) <> 0 Then
    InputStr = DecryptString(InputStr, UpdateKey)
    InputStr = Replace(InputStr, vbCrLf, "")
    Dim outputText As String
    outputText = InputStr
    ChallengeTemp.Show
    With ChallengeTemp.Dir1
        .Path = App.Path & "\Archive"
        For i = 1 To .ListCount
            If Dir(.list(i) & "\ChallengeLib") <> "" Then
                folderName = fso.GetFolder(.list(i)).Name
                inf = getPracticedCount(folderName, getTime, getCount)
                totalCount = totalCount + getCount
                totalTime = totalTime + getTime
            End If
        Next i
    End With
End If
rtnTime = totalTime
rtnCount = totalCount
outputText = outputText & Date & " " & Time & "::" & totalTime & "::" & totalCount & "\\"
    Set a = fs.CreateTextFile(App.Path & "\punchToday", True)
    a.WriteLine EncryptString(outputText, UpdateKey) '默认练习强度
    a.Close
    Set a = Nothing
    If ChallengeTemp.Visible = True Then Unload ChallengeTemp
End Function

Public Function getPunchRecords() As Long
    ' 文件不存在或为空时返回0
    Dim filePath As String
    filePath = App.Path & "\punchRecords"
    
    If Dir(filePath) = "" Then
        getPunchRecords = 0
        Exit Function
    End If
    
    ' 读取并解密文件内容
    OpenTxt filePath
    Dim punchData As String
    punchData = DecryptString(InputStr, UpdateKey)
    
    If Len(Replace(punchData, vbCrLf, "")) < 3 Then
        getPunchRecords = 0
        Exit Function
    End If
    
    ' 解析所有记录
    Dim block() As String, unit() As String
    block = Split(punchData, "\\")
    
    ' 今天日期
    Dim today As Date
    today = Date
    
    ' 收集有效打卡日期和类型
    Dim dates() As Date
    Dim isMakeup() As Boolean
    ReDim dates(UBound(block))
    ReDim isMakeup(UBound(block))
    
    Dim count As Long
    count = 0
    
    For i = LBound(block) To UBound(block) - 1
        unit = Split(block(i), "::")
        If UBound(unit) >= 3 Then
            dates(count) = CDate(unit(0))
            isMakeup(count) = (unit(3) = "mTrue")
            count = count + 1
        End If
    Next i
    
    If count = 0 Then
        getPunchRecords = 0
        Exit Function
    End If
    
    ' 按日期排序
'    Dim i As Long, j As Long
    For i = 0 To count - 2
        For j = i + 1 To count - 1
            If dates(i) > dates(j) Then
                Dim tempDate As Date
                tempDate = dates(i)
                dates(i) = dates(j)
                dates(j) = tempDate
                
                Dim tempFlag As Boolean
                tempFlag = isMakeup(i)
                isMakeup(i) = isMakeup(j)
                isMakeup(j) = tempFlag
            End If
        Next j
    Next i
    
    ' 核心逻辑：从最新记录往前计算连续天数
    Dim streak As Long
    streak = 1  ' 最新记录本身算1天
    
    ' 从倒数第二条记录开始往前检查
    For i = count - 2 To 0 Step -1
        Dim daysDiff As Long
        daysDiff = DateDiff("d", dates(i), dates(i + 1))
        
        If daysDiff = 1 Then
            ' 连续打卡
            streak = streak + 1
        ElseIf daysDiff > 1 Then
            ' 缺卡
            If isMakeup(i + 1) Then
                ' 当前记录是补卡，检查是否补上了缺卡
                If DateDiff("d", dates(i), dates(i + 1) - 1) <= 0 Then
                    ' 补卡日期正好是缺卡日，可以连接
                    streak = streak + 1
                Else
                    ' 补卡没有补上连续的缺卡，中断
                    Exit For
                End If
            Else
                ' 当前记录是正常打卡，检查缺卡是否超过3天
                If daysDiff <= 4 Then  ' 因为缺卡n天意味着间隔n+1天
                    ' 3天内的缺卡，检查中间是否有补卡
                    Dim hasMakeupInBetween As Boolean
                    hasMakeupInBetween = False
                    
                    For j = i + 1 To count - 1
                        If isMakeup(j) Then
                            Dim makeupDaysDiff As Long
                            makeupDaysDiff = DateDiff("d", dates(i), dates(j))
                            If makeupDaysDiff >= 1 And makeupDaysDiff <= 3 Then
                                hasMakeupInBetween = True
                                Exit For
                            End If
                        End If
                    Next j
                    
                    If hasMakeupInBetween Then
                        ' 有补卡，可以继续连接
                        streak = streak + 1
                    Else
                        ' 无补卡，中断
                        Exit For
                    End If
                Else
                    ' 缺卡超过3天，中断
                    Exit For
                End If
            End If
        End If
    Next i
    
    ' 检查最近一次打卡到今天是否超过3天
    If DateDiff("d", dates(count - 1), today) > 3 Then
        streak = 0
    End If
    
    getPunchRecords = streak
End Function


Public Function getLibStats(ByVal libname As String, ByRef pastWrong As Integer, ByRef cDifficulty As Integer, ByRef unPracticed As Integer)
pastWrong = 0
cDifficulty = 0
unPracticed = 0
Dim totalCount As Integer
totalCount = GetFileCount(App.Path & "\Archive\" & libname & "\topics")
If totalCount = 0 Then Exit Function
totalCount = Int(totalCount / 4)
pastWrong = getFaultcount(libname)


For i = 1 To totalCount
    OpenTxt (App.Path & "\Archive\" & libname & "\topics\" & i & "_stars")
    InputStr = Replace(InputStr, vbCrLf, "")
    If Val(InputStr) > 75 Then cDifficulty = cDifficulty + 1
    If Dir(App.Path & "\Archive\" & libname & "\PracticeRecords\" & i & "_records") = "" Then
        If Dir(App.Path & "\Archive\" & libname & "\topics\" & i & "_topic") <> "" Then '未被禁用，但确实没练过
            unPracticed = unPracticed + 1
        End If
    End If
Next i


End Function

Public Function getPracticedCount(libname As String, ByRef pTime As Integer, pCount As Integer)

                OpenTxt App.Path & "\Archive\" & libname & "\ChallengeLog"
                logline = Split(InputStr, vbCrLf)
                Dim startTime As String
                Dim endTime As String
                Dim minutesDiff As Integer
                Dim totalCount As Long
                Dim totalTime As Integer
                totalCount = 0
                totalTime = 0
                For i = 1 To UBound(logline)
                    If Len(logline(i)) > 6 Then
                        lineBlock = Split(logline(i), "::")
                        If lineBlock(3) = "1" Then
                            TargetTest = lineBlock(2)
                            startTime = lineBlock(0)
                            For j = i To UBound(logline)
                                If Len(logline(j)) > 6 Then
                                    lineBlockCheck = Split(logline(j), "::")
                                    If lineBlockCheck(2) = TargetTest And lineBlockCheck(3) = "2" Then
                                        'record end time
                                        username = Environ("USERNAME")
                                        foldercount = GetFolderCount("C:\Users\" & username & "\TestRecords\" & TargetTest)
                                        For kk = 1 To foldercount
                                            If Dir("C:\Users\" & username & "\TestRecords\" & TargetTest & "\" & kk & "\span") <> "" Then
                                                OpenTxt "C:\Users\" & username & "\TestRecords\" & TargetTest & "\" & kk & "\span"
                                                If Val(Replace(InputStr, vbCrLf, "")) <> 0 Then
                                                    totalCount = totalCount + 1
                                                    totalTime = totalTime + Int(Replace(InputStr, vbCrLf, ""))
                                                End If
                                            End If
                                        Next kk
'                                        endTime = lineBlockCheck(0)
'                                        minutesDiff = DateDiff("n", startTime, endTime)
'                                        totalTime = totalTime + minutesDiff
                                    End If
                                End If
                            Next j
                        End If
                    End If
                Next i
pCount = Int(totalCount)
pTime = Int(totalTime / 60)
End Function



Public Function GetShuffledArray(Optional ByVal minNum As Integer = 1, Optional ByVal maxNum As Integer) As Variant
    Dim i As Integer, j As Integer
    Dim temp As Integer
    Dim arr() As Integer
    Dim count As Integer
    
    count = maxNum - minNum + 1
    ReDim arr(1 To count)
    
    ' 初始化数组
    For i = 1 To count
        arr(i) = minNum + i - 1
    Next i
    
    ' 洗牌
    Randomize timer
    
    For i = count To 2 Step -1
        j = Int(Rnd * i) + 1
        
        ' 交换
        temp = arr(i)
        arr(i) = arr(j)
        arr(j) = temp
    Next i
    
    GetShuffledArray = arr
End Function

Public Function PatchLibTrainingInfo(libname As String)
If Dir(App.Path & "\Archive\" & libname & "\OptimizedProperties") = "" Or Dir(App.Path & "\Archive\" & libname & "\ChallengeLog") = "" Then '例行检查是否文件缺失，这样遍历所有挑战题库时都可以执行此功能，缺失自动补齐
    Dim a As Object, fs As Object
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set a = fs.CreateTextFile(App.Path & "\Archive\" & libname & "\ChallengeLib", True)
    a.WriteLine "1000" '默认练习强度
    a.Close
    Set a = Nothing
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set a = fs.CreateTextFile(App.Path & "\Archive\" & libname & "\ChallengeTasks", True)
    Dim shuffledNumbers As String
    shuffledNumbers = ShuffleNumbers()
    a.WriteLine "tasks;" & shuffledNumbers & vbCrLf & "1" & vbCrLf & "2"
    a.Close
    Set a = Nothing
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set a = fs.CreateTextFile(App.Path & "\Archive\" & libname & "\ChallengeLog", True)
    a.WriteLine Date
    a.Close
    Set a = Nothing

    'get indexPriority file
    Dim fileCount As Integer
    fileCount = GetFileCount(App.Path & "\Archive\" & libname & "\topics")
    Dim count As Integer
    count = 0
    Dim returnPriority As Double, returnLoad As Single
    outputResult = ""
    If fileCount <> 0 Then
    outputResult = libname & vbCrLf
        For k = 1 To fileCount / 4
            If Dir(App.Path & "\Archive\" & libname & "\topics\" & k & "_topic") <> "" Then
                count = count + 1
                msg = calculatePriorityScore(libname, Int(k), returnPriority, returnLoad)
                outputResult = outputResult & returnPriority & ":::" & returnLoad & vbCrLf
            Else
                outputResult = outputResult & "0" & vbCrLf
            End If
        Next k
    End If
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set a = fs.CreateTextFile(App.Path & "\Archive\" & libname & "\indexPriority", True)
    a.WriteLine outputResult
    a.Close
    Set a = Nothing
    
    Dim Result As String
    Result = "[ExamSource]=" & libname & vbCrLf
    Result = Result & "[ExamScalar]=20" & vbCrLf
    Result = Result & "[PointPerTopic]=10" & vbCrLf
    Result = Result & "[ThresholdValue]=0" & vbCrLf
    Result = Result & "[CorrectionMode]=Auto" & vbCrLf
    Result = Result & "[SetMins]=120" & vbCrLf
    Result = Result & "[Difficulty]=Easy" & vbCrLf
    Result = Result & "[AllowRedo]=True" & vbCrLf
    Result = Result & "[Status]=Unfinished"
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set a = fs.CreateTextFile(App.Path & "\Archive\" & libname & "\OptimizedProperties", True)
    a.WriteLine Result
    a.Close
    Set a = Nothing
    
    Result = ""
    
    Result = "[AllowPause] = True" & vbCrLf
    Result = Result & "[AllowLeave] = True" & vbCrLf
    Result = Result & "[AntiCheatService] = False" & vbCrLf
    Result = Result & "[Account] = False" & vbCrLf
    Result = Result & "InstantJudgeEnabled"
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set a = fs.CreateTextFile(App.Path & "\Archive\" & libname & "\OptimizedPropertiesEtc", True)
    a.WriteLine Result
    a.Close
    Set a = Nothing
    
    Result = ""
    
    Result = "Unavailable" & vbCrLf & "OneClickFalse" & vbCrLf & "ViewOnceFalse" & vbCrLf & "OptionShuffleTrue"
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set a = fs.CreateTextFile(App.Path & "\Archive\" & libname & "\OptSelectorOptions", True)
    a.WriteLine Result
    a.Close
    Set a = Nothing
End If
End Function



Public Function SortErrorList(ByVal OriginalList As String, ByVal SortBy As Integer) As String
    ' SortBy 参数说明：
    ' 1 = 按错误次数排序
    ' 2 = 按最近一次答错距今的天数排序
    
    Dim arrBlocks() As String
    Dim arrData() As String
    Dim i As Long, j As Long
    Dim temp As String
    Dim swapped As Boolean
    
    ' 使用 } 分割原始字符串
    arrBlocks = Split(OriginalList, "}")
    
    ' 冒泡排序实现（简单易懂）
    For i = LBound(arrBlocks) To UBound(arrBlocks) - 1
        swapped = False
        For j = LBound(arrBlocks) To UBound(arrBlocks) - i - 1
            If arrBlocks(j) <> "" And arrBlocks(j + 1) <> "" Then
                Dim block1() As String
                Dim block2() As String
                
                block1 = Split(arrBlocks(j), ":")
                block2 = Split(arrBlocks(j + 1), ":")
                
                ' 确保每个数据块都有足够的数据
                If UBound(block1) >= 2 And UBound(block2) >= 2 Then
                    Dim val1 As Long, val2 As Long
                    
                    ' 根据 SortBy 选择比较的值
                    If SortBy = 1 Then
                        val1 = CLng(block1(1))   ' 错误次数
                        val2 = CLng(block2(1))
                    ElseIf SortBy = 2 Then
                        val1 = CLng(block1(2))   ' 最近一次答错距今的天数
                        val2 = CLng(block2(2))
                    Else
                        ' 默认按错误次数排序
                        val1 = CLng(block1(1))
                        val2 = CLng(block2(1))
                    End If
                    
                    ' 如果前一个值小于后一个值，则交换（降序排序）
                    If val1 < val2 Then
                        temp = arrBlocks(j)
                        arrBlocks(j) = arrBlocks(j + 1)
                        arrBlocks(j + 1) = temp
                        swapped = True
                    End If
                End If
            End If
        Next j
        
        ' 如果本轮没有交换，说明已经排序完成
        If Not swapped Then Exit For
    Next i
    
    ' 重新构建排序后的字符串
    Dim Result As String
    Result = ""
    
    For i = LBound(arrBlocks) To UBound(arrBlocks)
        If arrBlocks(i) <> "" Then
            If Result <> "" Then
                Result = Result & arrBlocks(i) & "}"
            Else
                Result = arrBlocks(i) & "}"
            End If
        End If
    Next i
    
    ' 移除最后一个多余的 }（如果有的话）
    If Right(Result, 1) = "}" Then
        Result = Left(Result, Len(Result) - 1)
    End If
    
    SortErrorList = Result
End Function


Public Function GetFileCount(ByVal FolderPath As String, Optional ByVal fileFilter As String = "*.*") As Long
    Dim objFolder As Object
    Dim objFiles As Object
    Dim objFile As Object
    Dim count As Long
    
    ' 校验文件夹是否存在
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Dir(FolderPath, vbDirectory) = "" Then
        MsgBox "文件夹不存在：" & FolderPath, vbCritical
        GetFileCount = -1 ' 返回-1表示异常
        Exit Function
    End If
    
    ' 获取文件夹对象
    Set objFolder = fso.GetFolder(FolderPath)
    Set objFiles = objFolder.Files
    count = 0
    
    ' 遍历文件（支持过滤类型，如"*.txt" "*.zip"）
    For Each objFile In objFiles
        ' 匹配文件类型（不区分大小写）
            count = count + 1
            ' 可选：获取单个文件的详细信息
    Next objFile
    
    GetFileCount = count
    ' 释放对象
    Set objFile = Nothing
    Set objFiles = Nothing
    Set objFolder = Nothing
End Function

Function GetFolderCount(ByVal FolderPath As String) As Long
    Dim fso As Object
    Dim fld As Object
    Dim subFolder As Object
    Dim count As Long
    
    ' 创建 FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' 检查文件夹是否存在
    If Not fso.FolderExists(FolderPath) Then
        GetFolderCount = -1 ' 文件夹不存在
        Exit Function
    End If
    
    ' 获取文件夹对象
    Set fld = fso.GetFolder(FolderPath)
    
    ' 计算子文件夹数量
    count = 0
    For Each subFolder In fld.SubFolders
        count = count + 1
    Next subFolder
    
    GetFolderCount = count
    
    ' 清理对象
    Set subFolder = Nothing
    Set fld = Nothing
    Set fso = Nothing
End Function

Function HoursUntilTomorrow() As Double
    Dim currentTime As Date
    Dim tomorrowMidnight As Date
    Dim hoursRemaining As Double
    
    ' 获取当前时间
    currentTime = Now
    
    ' 计算明天凌晨0点的时间
    tomorrowMidnight = DateAdd("d", 1, DateValue(Now))
    
    ' 计算时间差（以天为单位），然后转换为小时
    hoursRemaining = (tomorrowMidnight - currentTime) * 24
    
    ' 返回结果
    HoursUntilTomorrow = hoursRemaining
End Function
Public Function checkTaskStats_readLog(inputNum As Integer)
OpenTxt App.Path & "\Archive\" & ChallengeLib & "\ChallengeLog"
logline = Split(InputStr, vbCrLf)
        For i = 1 To UBound(logline)
            If Len(logline(i)) > 6 Then
                lineBlock = Split(logline(i), "::")
                If Int(lineBlock(5)) = inputNum And lineBlock(3) = "1" Then
                    TargetTest = lineBlock(2)
                    For j = i To UBound(logline)
                        If Len(logline(j)) > 6 Then
                            lineBlockCheck = Split(logline(j), "::")
                            If lineBlockCheck(2) = TargetTest And lineBlockCheck(3) = "2" And Int(lineBlockCheck(5)) = inputNum Then
                                'missionAccomplished
                                taskAccomplished = True
                            End If
                        End If
                    Next j
                End If
            End If
        Next i


End Function
Public Function previousTestCheck(tasknum As Integer)
    OpenTxt App.Path & "\Archive\" & ChallengeLib & "\ChallengeLog"
    On Error Resume Next
    Dim logline() As String
    logline = Split(InputStr, vbCrLf)
    Dim lineBlock() As String
    taskAccomplished = False
    isGenerated = False
        For i = 1 To UBound(logline)
            If Len(logline(i)) > 6 Then
                lineBlock = Split(logline(i), "::")
                Dim TargetTest As String
                If Int(lineBlock(5)) = tasknum And lineBlock(3) = "1" Then
                    isGenerated = True
                    TargetTest = lineBlock(2)
                    Dim lineBlockCheck() As String
                    Dim notCompleted As Boolean
                    notCompleted = True
                    For j = i To UBound(logline)
                        If Len(logline(j)) > 6 Then
                            lineBlockCheck = Split(logline(j), "::")
                            If lineBlockCheck(2) = TargetTest And lineBlockCheck(3) = "2" And Int(lineBlockCheck(5)) = tasknum Then
                                'missionAccomplished
                                taskAccomplished = True
                                notCompleted = False
                            End If
                        End If
                    Next j
                    If notCompleted = True And isGenerated = True Then
                        '继续之前未完成的任务
                        username = Environ("USERNAME")
                        If Dir("C:\Users\" & username & "\TestRecords\" & TargetTest, vbDirectory) = "" Then
                            isGenerated = False
                        Else
                            startExercise (TargetTest)
                        End If
                        Exit Function
                    End If
                    GoTo checknext
                End If
            End If
checknext:
        Next i
        Debug.Print taskAccomplished
        Debug.Print notCompleted
        Debug.Print isGenerated
'        If isGenerated = False Then
'            target = ChallengeLib
'            IntelligentAgentGenerator.Show
'        End If

If isGenerated = True And taskAccomplished = True Then isGenerated = False


End Function
Public Function startExercise(TargetTest As String)

                        Exam_Foreplay.Show
                        Exam_Foreplay.Tag = TargetTest
                        username = Environ("USERNAME")
                        If Dir("C:\Users\" & username & "\TestRecords\" & TargetTest & "\index") <> "" Then
                            OpenTxt "C:\Users\" & username & "\TestRecords\" & TargetTest & "\index"
                            InputStr = DecryptString(InputStr, UpdateKey)
                            If InStr(InputStr, "CertifiedOnThisPC") = 0 Then
                                RunPassword.Show
                                passcode_correct = False
                                Dim b() As String
                                b = Split(InputStr, vbCrLf)
                                passcode_temp = b(1)
                            End If
                        End If
                        OpenTxt "C:\Users\" & username & "\TestRecords\" & TargetTest & "\properties"
                            Exam.Propinput.text = InputStr
                            If InStr(InputStr, "[AllowRedo]=False") <> 0 Then
                                Exam_Foreplay.Label3.Visible = True
                            Else
                                Exam_Foreplay.Label3.Visible = False
                            End If
                            InputStr = ""
                        If Dir("C:\Users\" & username & "\TestRecords\" & TargetTest & "\TimeLimited") <> "" Then
                        OpenTxt "C:\Users\" & username & "\TestRecords\" & TargetTest & "\TimeLimited"
                        InputStr = DecryptString(InputStr, UpdateKey)
                        Dim spliterA() As String, spliterB() As String, spliterC() As String
                        spliterA = Split(InputStr, vbCrLf)
                        spliterB = Split(spliterA(0), ";;;")
                        spliterC = Split(spliterA(1), ";;;")
                        If spliterA(1) <> "none" Then
                            t1 = Format(Date, "yyyy/mm/dd") & " " & Format(Time, "hh:mm:ss")
                            t2 = spliterC(0) & " " & spliterC(1)
                            h = Int(DateDiff("s", t1, t2) / 3600)
                        m = Int((DateDiff("s", t1, t2) - h * 3600) / 60)
                        s = DateDiff("s", t1, t2) - h * 3600 - m * 60
                        If h < 0 Or m < 0 Or s < 0 Then
                            '已超过截止日期，无效
                            Exam_Foreplay.Label2.Enabled = False
                            Exam_Foreplay.Label2.Caption = "已超过截止日期"
                            Exit Function
                        End If
                        End If
                            '检查是否在开始日期前
                        If spliterA(0) <> "none" Then
                            t1 = Format(Date, "yyyy/mm/dd") & " " & Format(Time, "hh:mm:ss")
                            t2 = spliterB(0) & " " & spliterB(1)
                            h = Int(DateDiff("s", t1, t2) / 3600)
                        m = Int((DateDiff("s", t1, t2) - h * 3600) / 60)
                        s = DateDiff("s", t1, t2) - h * 3600 - m * 60
                        If h > 0 And m > 0 And s > 0 Or h = 0 And m > 0 And s > 0 Or h = 0 And m = 0 And s > 0 Then
                            '未到开始时段，无效
                            Exam_Foreplay.Label2.Enabled = True
                            Exam_Foreplay.Label4.Visible = True
                            Exam_Foreplay.Label4.Caption = "距离考试开始还有" & h & "小时" & m & "分钟，您可以进入等候室等候。"
                            Exam_Foreplay.Label2.Caption = "进入等候室"
                            Exit Function
                        End If
                        If spliterA(1) <> "none" Then
                        t1 = Format(Date, "yyyy/mm/dd") & " " & Format(Time, "hh:mm:ss")
                            t2 = spliterC(0) & " " & spliterC(1)
                            h = Int(DateDiff("s", t1, t2) / 3600)
                            m = Int((DateDiff("s", t1, t2) - h * 3600) / 60)
                            s = DateDiff("s", t1, t2) - h * 3600 - m * 60
                            Exam_Foreplay.Label4.Visible = True
                            Exam_Foreplay.Label4.Caption = "距离截止时间还有" & h & "小时" & m & "分钟"
                        
                        
                        End If
                        End If
                        End If
End Function

Public Function checkLibRatingChange(ByRef oRating As Single, cRating As Single, mean As Single) As Boolean
    Dim logFile As String
    logFile = App.Path & "\Archive\" & ChallengeLib & "\ChallengeLog"
    
    If Dir(logFile) = "" Then Exit Function
    Dim logline() As String
    Dim block() As String
    Dim i As Long
    OpenTxt logFile ' 假设 OpenTxt 将文件内容读入 InputStr
    

    
    logline = Split(InputStr, vbCrLf)
    If UBound(logline) < 1 Then Exit Function ' 文件内容不足两行
    If DateDiff("d", logline(0), Date) <> 0 Then Exit Function '日期检查，必须当日
    block = Split(logline(1), "::") ' 第一行是初始记录
    If UBound(block) < 4 Then Exit Function ' 格式不符
    On Error Resume Next
    oRating = CSng(block(4)) ' 直接取数值，无需 Sin
    If Err.number <> 0 Then Exit Function
    On Error GoTo 0
    
    cRating = oRating ' 初始化为初始值
    
    For i = 1 To UBound(logline)
        If InStr(logline(i), "::") > 0 Then
            block = Split(logline(i), "::")
            If UBound(block) >= 4 Then
                On Error Resume Next
                cRating = CSng(block(4))
                If Err.number <> 0 Then
                    ' 转换失败，跳过该行
                Else
                    ' 成功获取当前值
                End If
                On Error GoTo 0
            End If
        End If
    Next i
    
    mean = cRating - oRating
    checkLibRatingChange = True ' 表示成功执行
End Function

Public Function CheckTaskStats()
On Error GoTo errDealer
If Dir(App.Path & "\Archive\Notes", vbDirectory) <> "" Then
    Dim FD As Object
    Set FD = CreateObject("Scripting.FileSystemObject")
    FD.DeleteFolder App.Path & "\Archive\Notes"
End If
If GetFileCount(App.Path & "\Archive\" & ChallengeLib & "\topics") = 0 Then
    msg = Dashboard.showCriticalMsg("该题库暂无题目！")
    Dashboard.FrameGtLib.Visible = True
    If ChallengeTemp.Visible = True Then Unload ChallengeTemp
    Exit Function
End If

OpenTxt App.Path & "\Archive\" & ChallengeLib & "\ChallengeLib"
Dim difficultyIndex As Integer
InputStr = Replace(InputStr, vbCrLf, "")
difficultyIndex = Int(InputStr)
If difficultyIndex = 0 Then
    MsgBox ("检查任务状态终止：缺少难度参数。")
    ChallengeDifficulty.Show
    Exit Function
End If
maxTopicNum = difficultyIndex
'Select Case difficultyIndex
'    Case 1
'        maxTopicNum = 10
'    Case 2
'        maxTopicNum = 20
'    Case 3
'        maxTopicNum = 30
'End Select
'Navigator_Main.FrameFinish.Visible = False
If ChallengeTaskNumberForToday = "" Then
    checkAvailableTasks
End If
CheckTaskNum = Split(ChallengeTaskNumberForToday, ";")

Dim checkOrder As Integer
checkOrder = 1
Dim faultFinished As Integer

Do Until checkOrder > 3
    taskAccomplished = False
    Select Case checkOrder
    Case 1
    OpenTxt App.Path & "\Archive\" & ChallengeLib & "\ChallengeLog"
    logline = Split(InputStr, vbCrLf)
        For i = 1 To UBound(logline)
            If Len(logline(i)) > 6 Then
                lineBlock = Split(logline(i), "::")
                Dim TargetTest As String
                If lineBlock(5) = "0" And lineBlock(3) = "1" Then
                    TargetTest = lineBlock(2)
                    Dim lineBlockCheck() As String
                    For j = i To UBound(logline)
                        If Len(logline(j)) > 6 Then
                            lineBlockCheck = Split(logline(j), "::")
                            If lineBlockCheck(2) = TargetTest And lineBlockCheck(3) = "2" Then
                                'missionAccomplished
                                taskAccomplished = True
                            End If
                        End If
                    Next j
                End If
            End If
        Next i
        If taskAccomplished = True Then
            With Dashboard
                .TaskActionBtn(0).Visible = False
                .ShapeTaskStat(0).BorderColor = &HC000&
'            Navigator_Main.BgTaskFinish(0).Visible = True
'            Navigator_Main.LabelFinish(0).Visible = True
            End With
        End If
    Case 2 To 3
        Select Case Int(CheckTaskNum(checkOrder - 2))
            Case 1 To 2
                checkTaskStats_readLog (Int(CheckTaskNum(checkOrder - 2)))
                If taskAccomplished = True Then
                    With Dashboard
                        .TaskActionBtn(checkOrder - 1).Visible = False
                        .ShapeTaskStat(checkOrder - 1).BorderColor = &HC000&
'                        Navigator_Main.BgTaskFinish(1).Visible = True
'                        Navigator_Main.LabelFinish(1).Visible = True
                    End With
                End If
            Case 3 Or 9
                faultFinished = 0
                Dim block() As String, unit() As String
                block = Split(getListofRecentFaults(ChallengeLib), "}")
                For i = LBound(block) To UBound(block)
                    If block(i) <> "" Then
                        unit = Split(block(i), ":")
                        Dim getScore As Double, getLoad As Single
                        msg = calculatePriorityScore(ChallengeLib, Int(unit(1)), getScore, getLoad)
                        If getScore = 0 Then '是错题，而且今天做过
                            If Dir(App.Path & "\Archive\" & ChallengeLib & "\topics\" & unit(1) & "_topic") <> "" Then '确保本题不是因为被禁用而导致score = 0
                                faultFinished = faultFinished + 1
                            End If
                        End If
                    End If
                Next i
                If faultFinished >= 10 Then taskAccomplished = True
            Case 4 To 6
                checkTaskStats_readLog (Int(CheckTaskNum(checkOrder - 2)))
                If taskAccomplished = True Then
                    With Dashboard
                        .TaskActionBtn(checkOrder - 1).Visible = False
                        .ShapeTaskStat(checkOrder - 1).BorderColor = &HC000&
'                        Navigator_Main.BgTaskFinish(1).Visible = True
'                        Navigator_Main.LabelFinish(1).Visible = True
                    End With
                End If
            Case 7
                taskAccomplished = False
                With Dashboard
                    .TaskActionBtn(checkOrder - 1).Visible = False
                    .ShapeTaskStat(checkOrder - 1).BorderColor = &HC000&
'                Navigator_Main.BgTaskFinish(1).Visible = False
'                Navigator_Main.LabelFinish(1).Visible = False
                End With
                OpenTxt App.Path & "\Archive\" & ChallengeLib & "\ChallengeLog"
                Dim initValue As Single
                Dim maxValue As Single
                logline = Split(InputStr, vbCrLf)
                For i = 1 To UBound(logline)
                    If Len(logline(i)) > 6 Then
                        lineBlock = Split(logline(i), "::")
                        If i = 1 Then initValue = lineBlock(4): maxValue = initValue
                        If lineBlock(4) >= maxValue Then maxValue = lineBlock(4)
                        
                    End If
                Next i
                Dim mean As Integer
                mean = (maxValue - initValue) * 100
                If mean >= 100 Then taskAccomplished = True
                If taskAccomplished = True Then
                    With Dashboard
                        .TaskActionBtn(checkOrder - 1).Visible = False
                        .ProgressBar(checkOrder - 2).Visible = False
                        .ShapeTaskStat(checkOrder - 1).BorderColor = &HC000&
                    End With
                Else
                    With Dashboard
                        .ProgressBar(checkOrder - 2).Visible = True
                        .ProgressBar(checkOrder - 2).Max = 100
                        .ProgressBar(checkOrder - 2).Value = mean
                        .TaskActionBtn(checkOrder - 1).Visible = False
                        .ShapeTaskStat(checkOrder - 1).BorderColor = &HE0E0E0
                    End With
                End If
            Case 8
                Dim totalTime As Integer
                OpenTxt App.Path & "\Archive\" & ChallengeLib & "\ChallengeLog"
                logline = Split(InputStr, vbCrLf)
                Dim startTime As String
                Dim endTime As String
                Dim minutesDiff As Integer
                Dim totalSec As Long
                totalSec = 0
                totalTime = 0
                For i = 1 To UBound(logline)
                    If Len(logline(i)) > 6 Then
                        lineBlock = Split(logline(i), "::")
                        If lineBlock(3) = "1" Then
                            TargetTest = lineBlock(2)
                            startTime = lineBlock(0)
                            For j = i To UBound(logline)
                                If Len(logline(j)) > 6 Then
                                    lineBlockCheck = Split(logline(j), "::")
                                    If lineBlockCheck(2) = TargetTest And lineBlockCheck(3) = "2" Then
                                        'record end time
                                        username = Environ("USERNAME")
                                        foldercount = GetFolderCount("C:\Users\" & username & "\TestRecords\" & TargetTest)
                                        For kk = 1 To foldercount
                                            If Dir("C:\Users\" & username & "\TestRecords\" & TargetTest & "\" & kk & "\span") <> "" Then
                                                OpenTxt "C:\Users\" & username & "\TestRecords\" & TargetTest & "\" & kk & "\span"
                                                totalSec = totalSec + Val(Replace(InputStr, vbCrLf, ""))
                                            End If
                                        Next kk
'                                        endTime = lineBlockCheck(0)
'                                        minutesDiff = DateDiff("n", startTime, endTime)
'                                        totalTime = totalTime + minutesDiff
                                    End If
                                End If
                            Next j
                        End If
                    End If
                Next i
                totalTime = Int(totalSec / 60)
                If totalTime >= 20 Then
                    taskAccomplished = True
                Else
                    taskAccomplished = False
                End If
                If taskAccomplished = True Then
                    With Dashboard
                        .TaskActionBtn(checkOrder - 1).Visible = False
                        .ProgressBar(checkOrder - 2).Visible = False
                        .ShapeTaskStat(checkOrder - 1).BorderColor = &HC000&
                    End With
                Else
                    With Dashboard
                        .TaskActionBtn(checkOrder - 1).Visible = False
                        .ProgressBar(checkOrder - 2).Visible = True
                        .ProgressBar(checkOrder - 2).Max = 20
                        .ProgressBar(checkOrder - 2).Value = totalTime
                    End With
                    
                End If
'            Case 9
'
'
'
'                OpenTxt App.Path & "\Archive\" & ChallengeLib & "\ChallengeLog"
'                logLine = Split(InputStr, vbCrLf)
'                Dim qualifiedTime As Integer
'                qualifiedTime = 0
'                For i = 1 To UBound(logLine)
'                    If Len(logLine(i)) > 6 Then
'                        lineBlock = Split(logLine(i), "::")
'                        If lineBlock(3) = "1" And Int(lineBlock(5)) = Int(CheckTaskNum(0)) Then
'                            TargetTest = lineBlock(2)
'                            startTime = lineBlock(0)
'                            For j = i To UBound(logLine)
'                                If Len(logLine(j)) > 6 Then
'                                    lineBlockCheck = Split(logLine(j), "::")
'                                    If lineBlockCheck(2) = TargetTest And lineBlockCheck(3) = "2" And Int(lineBlockCheck(5)) = Int(CheckTaskNum(0)) Then
'                                        'check correct
'                                        username = Environ("USERNAME")
'                                        OpenTxt ("C:\Users\" & username & "\TestRecords\" & TargetTest & "\Properties")
'                                        a = Split(InputStr, vbCrLf)
'                                        PointPerTopic = Split(a(2), "[PointPerTopic]=")
'                                        ExamScalar = Split(a(1), "[ExamScalar]=")
'                                        InputStr = ""
'                                        OpenTxt ("C:\Users\" & username & "\TestRecords\" & TargetTest & "\Conclusion")
'                                        targetScore = Val(InputStr)
'                                        If targetScore > Int(PointPerTopic) * Int(ExamScalar) * 0.8 Then
'                                            qualifiedTime = qualifiedTime + 1
'                                        End If
'
'                                    End If
'                                End If
'                            Next j
'                        End If
'                    End If
'                Next i
'                If qualifiedTime >= 2 Then
'                    taskAccomplished = True
'                Else
'                    taskAccomplished = False
'                End If
'                If taskAccomplished = True Then
'                    With Dashboard
'                        .TaskActionBtn(1).Visible = False
'                        .ShapeTaskStat(1).BorderColor = &HC000&
'                    End With
'                Else
'                    With Dashboard
'                        .TaskActionBtn(1).Visible = True
'                        .ProgressBar(0).Visible = True
'                        .ProgressBar(0).Max = 2
'                        .ProgressBar(0).Value = qualifiedTime
'                    End With
'                End If

        End Select
'    Case 3
'        Select Case Int(CheckTaskNum(1))
'            Case 1 To 6
'                checkTaskStats_readLog (Int(CheckTaskNum(1)))
'                If taskAccomplished = True Then
'                    With Dashboard
'                        .TaskActionBtn(2).Visible = False
'                        .ShapeTaskStat(2).BorderColor = &HC000&
'                    End With
'                End If
'            Case 7
'                taskAccomplished = False
'                With Dashboard
'                    .TaskActionBtn(2).Visible = True
'                End With
'                OpenTxt App.Path & "\Archive\" & ChallengeLib & "\ChallengeLog"
'                'Dim initValue As Single
'                'Dim maxValue As Single
'                logLine = Split(InputStr, vbCrLf)
'                For i = 1 To UBound(logLine)
'                    If Len(logLine(i)) > 6 Then
'                        lineBlock = Split(logLine(i), "::")
'                        If i = 1 Then initValue = lineBlock(4): maxValue = initValue
'                        If lineBlock(4) >= maxValue Then maxValue = lineBlock(4)
'
'                    End If
'                Next i
'                'Dim mean As Integer
'                mean = (maxValue - initValue) * 100
'                If mean >= 100 Then taskAccomplished = True
'                If taskAccomplished = True Then
'                    With Dashboard
'                            .TaskActionBtn(2).Visible = False
'                            .ProgressBar(1).Visible = False
'                            .ShapeTaskStat(2).BorderColor = &HC000&
'                    End With
'                Else
'                    With Dashboard
'                        .ProgressBar(1).Visible = True
'                        .ProgressBar(1).Max = 100
'                        .ProgressBar(1).Value = mean
'                        .TaskActionBtn(2).Visible = False
'                    End With
'                End If
'            Case 8
'                'Dim totalTime As Integer
'                OpenTxt App.Path & "\Archive\" & ChallengeLib & "\ChallengeLog"
'                logLine = Split(InputStr, vbCrLf)
'                'Dim startTime As String
'                'Dim endTime As String
'                'Dim minutesDiff As Integer
'                totalTime = 0
'                For i = 1 To UBound(logLine)
'                    If Len(logLine(i)) > 6 Then
'                        lineBlock = Split(logLine(i), "::")
'                        If lineBlock(3) = "1" Then
'                            TargetTest = lineBlock(2)
'                            startTime = lineBlock(0)
'                            For j = i To UBound(logLine)
'                                If Len(logLine(j)) > 6 Then
'                                    lineBlockCheck = Split(logLine(j), "::")
'                                    If lineBlockCheck(2) = TargetTest And lineBlockCheck(3) = "2" Then
'                                        'record end time
'                                        endTime = lineBlockCheck(0)
'                                        minutesDiff = DateDiff("n", startTime, endTime)
'                                        totalTime = totalTime + minutesDiff
'                                    End If
'                                End If
'                            Next j
'                        End If
'                    End If
'                Next i
'                If totalTime >= 20 Then
'                    taskAccomplished = True
'                Else
'                    taskAccomplished = False
'                End If
'                If taskAccomplished = True Then
'                    With Dashboard
'                        .TaskActionBtn(2).Visible = False
'                        .ProgressBar(1).Visible = False
'                        .ShapeTaskStat(2).BorderColor = &HC000&
'                    End With
'                Else
'                    With Dashboard
'                            .TaskActionBtn(2).Visible = False
'                            .ProgressBar(1).Visible = True
'                            .ProgressBar(1).Max = 20
'                            .ProgressBar(1).Value = totalTime
'                    End With
'                End If
'            Case 9
'
'                OpenTxt App.Path & "\Archive\" & ChallengeLib & "\ChallengeLog"
'                logLine = Split(InputStr, vbCrLf)
'                'Dim qualifiedTime As Integer
'                qualifiedTime = 0
'                For i = 1 To UBound(logLine)
'                    If Len(logLine(i)) > 6 Then
'                        lineBlock = Split(logLine(i), "::")
'                        If lineBlock(3) = "1" And Int(lineBlock(5)) = Int(CheckTaskNum(0)) Then
'                            TargetTest = lineBlock(2)
'                            startTime = lineBlock(0)
'                            For j = i To UBound(logLine)
'                                If Len(logLine(j)) > 6 Then
'                                    lineBlockCheck = Split(logLine(j), "::")
'                                    If lineBlockCheck(2) = TargetTest And lineBlockCheck(3) = "2" And Int(lineBlockCheck(5)) = Int(CheckTaskNum(0)) Then
'                                        'check correct
'                                        username = Environ("USERNAME")
'                                        OpenTxt ("C:\Users\" & username & "\TestRecords\" & TargetTest & "\Properties")
'                                        a = Split(InputStr, vbCrLf)
'                                        PointPerTopic = Split(a(2), "[PointPerTopic]=")
'                                        ExamScalar = Split(a(1), "[ExamScalar]=")
'                                        InputStr = ""
'                                        OpenTxt ("C:\Users\" & username & "\TestRecords\" & TargetTest & "\Conclusion")
'                                        targetScore = Val(InputStr)
'                                        If targetScore > Int(PointPerTopic) * Int(ExamScalar) * 0.8 Then
'                                            qualifiedTime = qualifiedTime + 1
'                                        End If
'
'                                    End If
'                                End If
'                            Next j
'                        End If
'                    End If
'                Next i
'                If qualifiedTime >= 2 Then
'                    taskAccomplished = True
'                Else
'                    taskAccomplished = False
'                End If
'                If taskAccomplished = True Then
'                    With Dashboard
'                            .TaskActionBtn(2).Visible = False
'                            .ShapeTaskStat(2).BorderColor = &HC000&
'                    End With
'                Else
'                    With Dashboard
'                            .TaskActionBtn(2).Visible = True
'                            .ProgressBar(1).Visible = True
'                            .ProgressBar(1).Max = 2
'                            .ProgressBar(1).Value = qualifiedTime
'                    End With
'                End If
'        End Select
    End Select
    
    
checkOrder = checkOrder + 1
Loop

Dim isEveryTaskAccomplished As Boolean
isEveryTaskAccomplished = True
For i = 0 To 2
    If Dashboard.ShapeTaskStat(i).BorderColor <> &HC000& Then
        isEveryTaskAccomplished = False
    End If
Next i
Exit Function
errDealer:
        Dim a As Object, fs As Object
        Set fs = CreateObject("Scripting.FileSystemObject")
        Set a = fs.CreateTextFile(App.Path & "\Archive\" & ChallengeLib & "\ChallengeTasks", True)
        Dim shuffledNumbers As String
        shuffledNumbers = ShuffleNumbers()
        a.WriteLine "tasks;" & shuffledNumbers
        a.Close
        Set a = Nothing
        checkAvailableTasks
'If isEveryTaskAccomplished = True Then
'    Dashboard.ShapeEx3.FillColor = &HF1FDD9
'    Dashboard.ShapeEx3.BorderColor = &HC0C000
'End If
End Function

Public Function RecordInitValue()
Dim a As Object, fs As Object
OpenTxt App.Path & "\Archive\" & ChallengeLib & "\ChallengeLog"
Dim lines() As String
lines = Split(InputStr, vbCrLf)
LogContent = ""
For i = LBound(lines) To UBound(lines)
    If Len(lines(i)) > 6 Then
        LogContent = LogContent & lines(i) & vbCrLf
    End If
Next i
Dim addLine As String
addLine = Time & "::" & target & "::eventRecordInitValue::-1"

'rating
ChallengeTemp.Show
DoEvents
ChallengeTemp.File1.Path = App.Path & "\Archive\" & ChallengeLib & "\topics"
Dim ratingLib() As Integer, n As Integer, low As Integer, midd As Integer, high As Integer, unPracticed As Integer
If ChallengeTemp.File1.ListCount <> 0 Then
    Dim Rating As Single
    Rating = 0
    Dim itemCount As Integer
    itemCount = 0
    For i = 1 To ChallengeTemp.File1.ListCount / 4
        If Dir(App.Path & "\Archive\" & ChallengeLib & "\topics\" & i & "_topic") <> "" Then
        itemCount = itemCount + 1
        End If
    Next i
    ReDim ratingLib(1 To ChallengeTemp.File1.ListCount / 4)
    For k = 1 To ChallengeTemp.File1.ListCount / 4
        If Dir(App.Path & "\Archive\" & ChallengeLib & "\topics\" & k & "_topic") <> "" Then
        OpenTxt App.Path & "\Archive\" & ChallengeLib & "\topics\" & k & "_stars"
        InputStr = Replace(InputStr, vbCrLf, "")
        star = Int(InputStr)
        ratingLib(k) = star
        If Dir(App.Path & "\Archive\" & ChallengeLib & "\PracticeRecords\" & k & "_records") <> "" Then
            If ratingLib(k) < 41 Then
                low = low + 1
            ElseIf ratingLib(k) >= 41 And ratingLib(k) < 75 Then
                midd = midd + 1
            ElseIf ratingLib(k) >= 75 Then
                high = high + 1
            End If
        Else
            unPracticed = unPracticed + 1
        End If
        Rating = Rating + star
        End If
    Next k
    totalrating = Rating
    avrRating = totalrating / itemCount

    
    'draw chart

    Rating = Rating / (itemCount)
    Rating = 100 - Rating
    Rating = (Rating / 80) * 100
    Rating = Format(Rating, "0.00")
Unload ChallengeTemp
addLine = addLine & "::" & Rating & "::-1"
LogContent = LogContent & addLine
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set a = fs.CreateTextFile(App.Path & "\Archive\" & ChallengeLib & "\ChallengeLog", True)
    a.WriteLine LogContent
    a.Close
    Set a = Nothing
End If




End Function



Public Function writeLog(ByRef currentItem As String, Action As Integer, taskCaseNum As String)
'currentItem:当前练习名称
'Action:当前的动作代码：
'0 - 无需动作代码
'1 - 开始
'2 - 结束
'3 - 暂停
'4 - 达成条件
'5 - 未达成条件
'数据组成：Time,target,currentitem,action,rating
checkLogValidity
Dim a As Object, fs As Object
OpenTxt App.Path & "\Archive\" & ChallengeLib & "\ChallengeLog"
Dim lines() As String
lines = Split(InputStr, vbCrLf)
LogContent = ""
For i = LBound(lines) To UBound(lines)
    If Len(lines(i)) > 6 Then
        LogContent = LogContent & lines(i) & vbCrLf
    End If
Next i
Dim addLine As String
addLine = Time & "::" & target & "::" & currentItem & "::" & Action

'rating
ChallengeTemp.Show
DoEvents
ChallengeTemp.File1.Path = App.Path & "\Archive\" & ChallengeLib & "\topics"
Dim ratingLib() As Integer, n As Integer, low As Integer, midd As Integer, high As Integer, unPracticed As Integer
If ChallengeTemp.File1.ListCount <> 0 Then
    Dim Rating As Single
    Rating = 0
    Dim itemCount As Integer
    itemCount = 0
    For i = 1 To ChallengeTemp.File1.ListCount / 4
        If Dir(App.Path & "\Archive\" & ChallengeLib & "\topics\" & i & "_topic") <> "" Then
        itemCount = itemCount + 1
        End If
    Next i
    ReDim ratingLib(1 To ChallengeTemp.File1.ListCount / 4)
    For k = 1 To ChallengeTemp.File1.ListCount / 4
        If Dir(App.Path & "\Archive\" & ChallengeLib & "\topics\" & k & "_topic") <> "" Then
        OpenTxt App.Path & "\Archive\" & ChallengeLib & "\topics\" & k & "_stars"
        InputStr = Replace(InputStr, vbCrLf, "")
        star = Int(InputStr)
        ratingLib(k) = star
        If Dir(App.Path & "\Archive\" & ChallengeLib & "\PracticeRecords\" & k & "_records") <> "" Then
            If ratingLib(k) < 41 Then
                low = low + 1
            ElseIf ratingLib(k) >= 41 And ratingLib(k) < 75 Then
                midd = midd + 1
            ElseIf ratingLib(k) >= 75 Then
                high = high + 1
            End If
        Else
            unPracticed = unPracticed + 1
        End If
        Rating = Rating + star
        End If
    Next k
    totalrating = Rating
    avrRating = totalrating / itemCount

    
    'draw chart

    Rating = Rating / (itemCount)
    Rating = 100 - Rating
    Rating = (Rating / 80) * 100
    Rating = Format(Rating, "0.00")
Unload ChallengeTemp
addLine = addLine & "::" & Rating & "::"
addLine = addLine & taskCaseNum
LogContent = LogContent & addLine
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set a = fs.CreateTextFile(App.Path & "\Archive\" & ChallengeLib & "\ChallengeLog", True)
    a.WriteLine LogContent
    a.Close
    Set a = Nothing
End If
End Function







Public Function GenerateExerciseFromTask(tasknum As Integer, trainCount As Integer)
username = Environ("USERNAME")
'On Error Resume Next
'OpenTxt App.Path & "\Archive\" & ChallengeLib & "\ChallengeLib"
'Dim difficultyIndex As Integer
'InputStr = Replace(InputStr, vbCrLf, "")
'difficultyIndex = Int(InputStr)
'If difficultyIndex = 0 Then
'    MsgBox ("从任务生成练习过程终止：缺少难度参数。")
'    ChallengeDifficulty.Show
'    Exit Function
'End If
'Select Case difficultyIndex
'    Case 1
'        maxTopicNum = 10
'    Case 2
'        maxTopicNum = 20
'    Case 3
'        maxTopicNum = 30
'End Select
'OpenTxt App.Path & "\ExamTemplate\" & Me.Tag & "\etc"
'If InStr(InputStr, "[Account] = True") Then
'    If account_name = "" Or account_class = "" Or account_authencode = "" Then
'        Account.Show
'        Account.Label15.Visible = True
'        Unload Me
'        Exit Sub
'    End If
'End If
InputStr = ""
If Dir(App.Path & "\Archive\" & ChallengeLib & "\OptimizedProperties") = "" Then
    PatchLibTrainingInfo (ChallengeLib)
End If
OpenTxt App.Path & "\Archive\" & ChallengeLib & "\OptimizedProperties"
Dim outputStr As String
outputStr = Replace(InputStr, "[ExamScalar]=20", "[ExamScalar]=" & trainCount)

'需要修改properties文件：题目数量
Dim f() As String, d() As String, e() As String, IsOptimized As Boolean
f = Split(InputStr, vbCrLf)
d = Split(f(0), "[ExamSource]=")
e = Split(f(1), "[ExamScalar]=")
If InStr(InputStr, "Optimized") <> 0 Then IsOptimized = True
Source = ChallengeLib
quantity = trainCount
i = 1
Do While Dir("C:\Users\" & username & "\TestRecords\everydayChallengeEvent++" & ChallengeLib & i, vbDirectory) <> ""
i = i + 1
Loop
MkDir ("C:\Users\" & username & "\TestRecords\everydayChallengeEvent++" & ChallengeLib & i)
Dim deploytarget As String
deploytarget = "everydayChallengeEvent++" & ChallengeLib & i
mm = writeLog(deploytarget, 1, str(tasknum))
Dim a As Object, fs As Object
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set a = fs.CreateTextFile("C:\Users\" & username & "\TestRecords\" & deploytarget & "\Properties", True)
    a.WriteLine outputStr
    a.Close
'initiate summon process
'preparation
Dim m As Long
Dim variable As Integer
Dim Sheet As Object, Stat As Object, Label As Integer, times As Single
For i = 1 To quantity
    MkDir ("C:\Users\" & username & "\TestRecords\" & deploytarget & "\" & i)
Next i
'处理resultPool
Dim block() As String, unit() As String
block = Split(resultPool, "}}")

For q = 1 To Val(quantity)
        '根据resultpool复制题目
        unit = Split(block(q - 1), "::")
        FileCopy App.Path & "\Archive\" & unit(0) & "\topics\" & unit(1) & "_topic", "C:\Users\" & username & "\TestRecords\" & deploytarget & "\" & q & "\topic"
        FileCopy App.Path & "\Archive\" & unit(0) & "\resolve\" & unit(1) & "_resolve", "C:\Users\" & username & "\TestRecords\" & deploytarget & "\" & q & "\resolve"
        FileCopy App.Path & "\Archive\" & unit(0) & "\answers\" & unit(1) & "_answer", "C:\Users\" & username & "\TestRecords\" & deploytarget & "\" & q & "\answer"
        FileCopy App.Path & "\Archive\" & unit(0) & "\sn", "C:\Users\" & username & "\TestRecords\" & deploytarget & "\source_sn"
        FileCopy App.Path & "\Archive\" & unit(0) & "\topics\" & unit(1) & "_sn", "C:\Users\" & username & "\TestRecords\" & deploytarget & "\" & q & "\topic_sn"
        If Dir(App.Path & "\Archive\" & unit(0) & "\topics\" & unit(1) & "_attachment", vbDirectory) <> "" Then
            MkDir ("C:\Users\" & username & "\TestRecords\" & deploytarget & "\" & q & "\attachment")
            Set fso = CreateObject("Scripting.FileSystemObject")
            fso.CopyFolder App.Path & "\Archive\" & unit(0) & "\topics\" & unit(1) & "_attachment", "C:\Users\" & username & "\TestRecords\" & deploytarget & "\" & q & "\attachment", True
        End If
        OpenTxt App.Path & "\Archive\" & unit(0) & "\topics\" & unit(1) & "_sn"
        InputStr = Replace(InputStr, vbCrLf, "")
        Dim information As String
        information = unit(0) & ":::" & unit(1) & ":::" & InputStr
        Set fs = CreateObject("Scripting.FileSystemObject")
        Set infoRecord = fs.CreateTextFile("C:\Users\" & username & "\TestRecords\" & deploytarget & "\" & q & "\sourceInfo", True)
        infoRecord.WriteLine information
        infoRecord.Close
        information = ""
        InputStr = ""


        Set fs = CreateObject("Scripting.FileSystemObject")
        Set Sheet = fs.CreateTextFile("C:\Users\" & username & "\TestRecords\" & deploytarget & "\" & q & "\sheet", True)
        Set Span = fs.CreateTextFile("C:\Users\" & username & "\TestRecords\" & deploytarget & "\" & q & "\span", True)
    Next q
'wipe out all labels or modify config
'File1.Path = App.Path & "\Archive\" & Source & "\topics"





FileCopy App.Path & "\Archive\" & ChallengeLib & "\OptimizedPropertiesEtc", "C:\Users\" & username & "\TestRecords\" & deploytarget & "\etc"
If Dir(App.Path & "\Archive\" & ChallengeLib & "\OptSelectorOptions") <> "" Then
    FileCopy App.Path & "\Archive\" & ChallengeLib & "\OptSelectorOptions", "C:\Users\" & username & "\TestRecords\" & deploytarget & "\SelectorOptions"
End If



'load exam entry
Exam_Foreplay.Show
Exam_Foreplay.Tag = deploytarget
OpenTxt "C:\Users\" & username & "\TestRecords\" & deploytarget & "\properties"
    Exam.Propinput.text = InputStr
    If InStr(InputStr, "[AllowRedo]=False") <> 0 Then
        Exam_Foreplay.Label3.Visible = True
    Else
        Exam_Foreplay.Label3.Visible = False
    End If
    InputStr = ""
Unload ChallengeTemp
End Function


Public Function ShuffleNumbers() As String
    Dim numbers(1 To 9) As Integer
    Dim i As Integer, j As Integer
    Dim temp As Integer
    Dim Result As String
    
    ' 初始化数组
    For i = 1 To 9
        numbers(i) = i
    Next i
    
    ' 使用Fisher-Yates算法打乱数组
    Randomize timer ' 初始化随机数生成器
    For i = 9 To 2 Step -1
        j = Int(Rnd * i) + 1 ' 生成1到i之间的随机数
        ' 交换元素
        temp = numbers(i)
        numbers(i) = numbers(j)
        numbers(j) = temp
    Next i
    
    ' 构建结果字符串
    For i = 1 To 9
        Result = Result & CStr(numbers(i))
        If i < 9 Then
            Result = Result & ";"
        End If
    Next i
    
    ShuffleNumbers = Result
End Function

Public Function shuffleSetCount()
    Dim shuffledItem() As String
    shuffledItem = Split(resultPool, "}}")
    itemCount = UBound(shuffledItem)
    Randomize timer ' 初始化随机数生成器
    Dim temp As String
    For i = UBound(shuffledItem) - 1 To LBound(shuffledItem) Step -1
        j = Int((i - LBound(shuffledItem)) * Rnd + LBound(shuffledItem))
        ' 交换元素
        temp = shuffledItem(i)
        shuffledItem(i) = shuffledItem(j)
        shuffledItem(j) = temp
    Next i
    resultPool = ""
    For i = LBound(shuffledItem) To UBound(shuffledItem)
        resultPool = resultPool & shuffledItem(i) & "}}"
    Next i
    Debug.Print resultPool
    
    OpenTxt App.Path & "\Archive\" & ChallengeLib & "\ChallengeLib"
    Dim difficultyIndex As Integer
    InputStr = Replace(InputStr, vbCrLf, "")
    difficultyIndex = Int(InputStr)
    If difficultyIndex = 0 Then
        MsgBox ("检查可用任务终止：缺少难度参数。")
        ChallengeDifficulty.Show
        Exit Function
    End If
    Select Case difficultyIndex
        Case 1
            maxTopicNum = 10
        Case 2
            maxTopicNum = 20
        Case 3
            maxTopicNum = 30
    End Select
    
    Select Case itemCount
        Case 0 To maxTopicNum
            ChallengeTopicMaximum = itemCount
        Case Else
            ChallengeTopicMaximum = maxTopicNum
    End Select



End Function

Public Function checkAvailableTasks()
If Dir(App.Path & "\Archive\Notes", vbDirectory) <> "" Then
    Dim FD As Object
    Set FD = CreateObject("Scripting.FileSystemObject")
    FD.DeleteFolder App.Path & "\Archive\Notes"
End If
If Dir(App.Path & "\Archive\" & ChallengeLib & "\ChallengeLib") = "" Then
    If Dir(App.Path & "\ChallengeLib") <> "" Then Kill App.Path & "\ChallengeLib"
    Exit Function
End If
OpenTxt App.Path & "\Archive\" & ChallengeLib & "\ChallengeLib"
Dim difficultyIndex As Integer
InputStr = Replace(InputStr, vbCrLf, "")
If InputStr = ChallengeLib Then difficultyIndex = 1 '某个意外错误导致存储difficultyIndex的存档显示值为Challengelib, 此为修复办法。

difficultyIndex = Int(InputStr)
If difficultyIndex = 0 Then
    MsgBox ("检查可用任务终止：缺少难度参数。")
    ChallengeDifficulty.Show
    Exit Function
End If
Select Case difficultyIndex
    Case 1
        maxTopicNum = 10
    Case 2
        maxTopicNum = 20
    Case 3
        maxTopicNum = 30
End Select
ChallengeTaskNumberForToday = ""
If Dir(App.Path & "\Archive\" & ChallengeLib & "\ChallengeTasks") <> "" Then
    '查看是否为今天，如果不是，则刷新再继续
    Dim checkDate As Date
    checkDate = FileDateTime(App.Path & "\Archive\" & ChallengeLib & "\ChallengeTasks")
    Dim cq() As String
    cq = Split(checkDate, " ")
    If cq(0) <> Date Then
        'refresh
        Dim a As Object, fs As Object
        Set fs = CreateObject("Scripting.FileSystemObject")
        Set a = fs.CreateTextFile(App.Path & "\Archive\" & ChallengeLib & "\ChallengeTasks", True)
        Dim shuffledNumbers As String
        shuffledNumbers = ShuffleNumbers()
        a.WriteLine "tasks;" & shuffledNumbers
        a.Close
        Set a = Nothing
    Else
        OpenTxt App.Path & "\Archive\" & ChallengeLib & "\ChallengeTasks"
        Dim cut() As String
        cut = Split(InputStr, vbCrLf)
        If UBound(cut) >= 2 Then
            ChallengeTaskNumberForToday = cut(1) & ";" & cut(2) & ";"
            GoTo Final
        End If
    End If
    '结束检查
    OpenTxt (App.Path & "\Archive\" & ChallengeLib & "\ChallengeTasks")
    InputStr = Replace(InputStr, vbCrLf, "")
    Dim task() As String
    task = Split(InputStr, ";")
    taskMaximum = 2
    taskCurrentCount = 0
    Dim fetchFaults As Integer, fetchHard As Integer, fetchUnpracticed As Integer
    msg = getLibStats(ChallengeLib, fetchFaults, fetchHard, fetchUnpracticed)
    Dim previousDate As Date
    Dim items() As String, block() As String, itemCount As Integer
    For i = 1 To 9
        If task(i) <> "0" And taskCurrentCount < 2 Then
            '这里检查一下这个任务能不能做
            Dim taskAvailable As Boolean
            taskAvailable = False
            Select Case Int(task(i))
                Case 1
                    taskAvailable = True
                Case 2
                    taskAvailable = True
                Case 3
                    If fetchFaults > 20 Then taskAvailable = True
                    
'                    previousDate = DateAdd("d", -20, Date)
'                    Conditions = "Lib:::" & ChallengeLib & ":::}}}DTModified:::" & previousDate & ":::after:::lastPracticed"
'                    aa = challenge_searchEvents
'                    '查询resultPool中有效的，并提纯
'                    items = Split(resultPool, "}}")
'                    itemCount = UBound(items)
'                    For ii = LBound(items) To UBound(items) - 1
'                        block = Split(items(ii), "::")
'                        If Dir(App.Path & "\Archive\" & ChallengeLib & "\PracticeRecords\" & block(1) & "_records") <> "" Then
'                            '是否为近期答错，而不是曾经答错
'                            OpenTxt App.Path & "\Archive\" & ChallengeLib & "\PracticeRecords\" & block(1) & "_records"
'                            Dim recordItem() As String, recordBlock() As String
'                            Dim currentValidity As Boolean
'                            currentVaildity = False
'                            recordItem = Split(InputStr, vbCrLf)
'                            For m = LBound(recordItem) To UBound(recordItem)
'                                If Len(recordItem(m)) > 5 Then
'                                    recordBlock = Split(recordItem(m), ":::")
'                                    If DateDiff("d", previousDate, recordBlock(0)) >= 0 And recordBlock(2) = "false" Then
'                                    '通过智能体评判算法，判断这题今天有没有做过，如果做过则为0，应删除。
'                                        Dim getScore As Double, getLoad As Single
'                                        msg = calculatePriorityScore(ChallengeLib, Int(block(1)), getScore, getLoad)
'                                        If getScore <> 0 Then
'                                        currentValidity = True
'                                        End If
'                                    End If
'                                End If
'                            Next m
'
'                            If currentValidity = False Then
'                                resultPool = Replace(resultPool, items(ii) & "}}", "")
'                                itemCount = itemCount - 1
'                            End If
'                        End If
'                    Next ii
'                    If itemCount > 9 Then taskAvailable = True
                Case 4
'                    previousDate = DateAdd("d", -7, Date)
'                    Conditions = "Lib:::" & ChallengeLib & ":::}}}DTModified:::" & previousDate & ":::earlier:::lastPracticed}}}"
'                    items = Split(resultPool, "}}")
'                    itemCount = UBound(items)
                    Dim countUnPrac As Integer
                    countUnPrac = fetchUnpracticed
                    If countUnPrac <= 20 And countUnPrac > 0 Then taskAvailable = True
'                    If itemCount > 9 Then taskAvailable = True
                Case 5, 10
                    ChallengeTemp.Show
                    DoEvents
                    ChallengeTemp.File1.Path = App.Path & "\Archive\" & ChallengeLib & "\topics"
                    Dim ratingLib() As Integer, n As Integer, low As Integer, midd As Integer, high As Integer, unPracticed As Integer
                    If ChallengeTemp.File1.ListCount <> 0 Then
                        Dim Rating As Single
                        Rating = 0
                        itemCount = 0
                        For ii = 1 To ChallengeTemp.File1.ListCount / 4
                            If Dir(App.Path & "\Archive\" & ChallengeLib & "\topics\" & ii & "_topic") <> "" Then
                            itemCount = itemCount + 1
                            End If
                        Next ii
                        ReDim ratingLib(1 To ChallengeTemp.File1.ListCount / 4)
                        For k = 1 To ChallengeTemp.File1.ListCount / 4
                            If Dir(App.Path & "\Archive\" & ChallengeLib & "\topics\" & k & "_topic") <> "" Then
                            OpenTxt App.Path & "\Archive\" & ChallengeLib & "\topics\" & k & "_stars"
                            InputStr = Replace(InputStr, vbCrLf, "")
                            star = Int(InputStr)
                            ratingLib(k) = star
                            If Dir(App.Path & "\Archive\" & ChallengeLib & "\PracticeRecords\" & k & "_records") <> "" Then
                                If ratingLib(k) < 41 Then
                                    low = low + 1
                                ElseIf ratingLib(k) >= 41 And ratingLib(k) < 75 Then
                                    midd = midd + 1
                                ElseIf ratingLib(k) >= 75 Then
                                    high = high + 1
                                End If
                            Else
                                unPracticed = unPracticed + 1
                            End If
                            Rating = Rating + star
                            End If
                        Next k
                        totalrating = Rating
                        avrRating = totalrating / itemCount
                    '遍历查找符合条件的题目
                    resultPool = ""
                    itemCount = 0
                        For k = 1 To ChallengeTemp.File1.ListCount / 4
                            If Dir(App.Path & "\Archive\" & ChallengeLib & "\topics\" & k & "_topic") <> "" Then
                                OpenTxt App.Path & "\Archive\" & ChallengeLib & "\topics\" & k & "_stars"
                                InputStr = Replace(InputStr, vbCrLf, "")
                                star = Int(InputStr)
                                Dim getScore As Double, getLoad As Single
                                msg = calculatePriorityScore(ChallengeLib, Int(k), getScore, getLoad)
                                If star >= avrRating And getScore <> 0 Then itemCount = itemCount + 1
                            
                            End If
                        Next k
                    End If
                        If itemCount > 14 Then taskAvailable = True
                Case 6
                    ChallengeTemp.Show
                    DoEvents
                    ChallengeTemp.File1.Path = App.Path & "\Archive\" & ChallengeLib & "\topics"
                    resultPool = ""
                    If ChallengeTemp.File1.ListCount <> 0 Then
                        Rating = 0
                        itemCount = 0
                    '    For i = 1 To ChallengeTemp.File1.ListCount / 4
                    '        If Dir(App.Path & "\Archive\" & ChallengeLib & "\topics\" & i & "_topic") <> "" Then
                    '        itemCount = itemCount + 1
                    '        End If
                    '    Next i
                        ReDim ratingLib(1 To ChallengeTemp.File1.ListCount / 4)
                        For k = 1 To ChallengeTemp.File1.ListCount / 4
                            If Dir(App.Path & "\Archive\" & ChallengeLib & "\topics\" & k & "_topic") <> "" Then
                            OpenTxt App.Path & "\Archive\" & ChallengeLib & "\topics\" & k & "_stars"
                            InputStr = Replace(InputStr, vbCrLf, "")
                            star = Int(InputStr)
                            ratingLib(k) = star
                            If Dir(App.Path & "\Archive\" & ChallengeLib & "\PracticeRecords\" & k & "_records") <> "" Then
                                msg = calculatePriorityScore(ChallengeLib, Int(k), getScore, getLoad)
                                If ratingLib(k) >= 75 And getScore <> 0 Then
                                    itemCount = itemCount + 1
                                End If
                            Else
                                unPracticed = unPracticed + 1
                            End If
                            Rating = Rating + star
                            End If
                        Next k
                    '    totalRating = Rating
                    '    avrRating = totalRating / itemCount
                    '遍历查找符合条件的题目
                        If itemCount > 9 Then taskAvailable = True
                    End If
                Case 7
                    taskAvailable = True
                Case 8
                    taskAvailable = True
                Case 9
                    If fetchFaults > 20 Then taskAvailable = True
'                    previousDate = DateAdd("d", -20, Date)
'                    Conditions = "Lib:::" & ChallengeLib & ":::}}}DTModified:::" & previousDate & ":::after:::lastPracticed"
'                    aa = challenge_searchEvents
'                    '查询resultPool中有效的，并提纯
'                    items = Split(resultPool, "}}")
'                    itemCount = UBound(items)
'                    For ii = LBound(items) To UBound(items) - 1
'                        block = Split(items(ii), "::")
'                        If Dir(App.Path & "\Archive\" & ChallengeLib & "\PracticeRecords\" & block(1) & "_records") <> "" Then
'                            '是否为近期答错，而不是曾经答错
'                            OpenTxt App.Path & "\Archive\" & ChallengeLib & "\PracticeRecords\" & block(1) & "_records"
'                            currentVaildity = False
'                            recordItem = Split(InputStr, vbCrLf)
'                            For m = LBound(recordItem) To UBound(recordItem)
'                                If Len(recordItem(m)) > 5 Then
'                                    recordBlock = Split(recordItem(m), ":::")
'                                    If DateDiff("d", previousDate, recordBlock(0)) >= 0 And recordBlock(2) = "false" Then
'                                    '通过智能体评判算法，判断这题今天有没有做过，如果做过则为0，应删除。
'
'                                        msg = calculatePriorityScore(ChallengeLib, Int(block(1)), getScore, getLoad)
'                                        If getScore <> 0 Then
'                                        currentValidity = True
'                                        End If
'                                    End If
'                                End If
'                            Next m
'
'                            If currentValidity = False Then
'                                resultPool = Replace(resultPool, items(ii) & "}}", "")
'                                itemCount = itemCount - 1
'                            End If
'                        End If
'                    Next ii
'                    If itemCount > maxTopicNum Then taskAvailable = True

            End Select
            If taskAvailable = True Then
                ChallengeTaskNumberForToday = ChallengeTaskNumberForToday & task(i) & ";"
                taskCurrentCount = taskCurrentCount + 1
            End If
        End If
    Next i

Debug.Print ChallengeTaskNumberForToday
'bb = writeLog("", 0, "TaskGenerated:" & ChallengeTaskNumberForToday)

'If cq(0) <> Date Then
        OpenTxt App.Path & "\Archive\" & ChallengeLib & "\ChallengeTasks"
        InputStr = Replace(InputStr, vbCrLf, "")
        Dim Digit() As String
        Digit = Split(ChallengeTaskNumberForToday, ";")
        Set fs = CreateObject("Scripting.FileSystemObject")
        Set a = fs.CreateTextFile(App.Path & "\Archive\" & ChallengeLib & "\ChallengeTasks", True)
        a.WriteLine InputStr & vbCrLf & Digit(0) & vbCrLf & Digit(1)
        a.Close
        Set a = Nothing

'End If
If ChallengeTemp.Visible = True Then Unload ChallengeTemp
Final:
Dim taskblock() As String
taskblock = Split(ChallengeTaskNumberForToday, ";")
Dim TaskCardIndex As Integer
For i = 0 To 1
    TaskCardIndex = i + 1
    Select Case Int(taskblock(i))
        Case 1
            Dashboard.LabelTaskName(TaskCardIndex).Caption = "练习准确率总体较低的题目"
'            Navigator_Main.PictureIcon(TaskCardIndex).Picture = Navigator_Main.ImageList2.ListImages(4).Picture
            Dashboard.TaskActionBtn(TaskCardIndex).Tag = "1"
        Case 2
            Dashboard.LabelTaskName(TaskCardIndex).Caption = "完成一次随机练习"
'            Navigator_Main.PictureIcon(TaskCardIndex).Picture = Navigator_Main.ImageList2.ListImages(8).Picture
            Dashboard.TaskActionBtn(TaskCardIndex).Tag = "2"
            Dashboard.ProgressBar(i).Visible = False
        Case 3
            Dashboard.LabelTaskName(TaskCardIndex).Caption = "复习近期的错题"
'            Navigator_Main.PictureIcon(TaskCardIndex).Picture = Navigator_Main.ImageList2.ListImages(7).Picture
            Dashboard.ProgressBar(i).Visible = True
            Dashboard.ProgressBar(i).Max = 10
            Dashboard.TaskActionBtn(TaskCardIndex).Tag = "3"
            Dashboard.TaskActionBtn(TaskCardIndex).Visible = True
        Case 4
            Dashboard.LabelTaskName(TaskCardIndex).Caption = "做完没有练习的题目"
'            Navigator_Main.PictureIcon(TaskCardIndex).Picture = Navigator_Main.ImageList2.ListImages(6).Picture
            Dashboard.TaskActionBtn(TaskCardIndex).Tag = "4"
            Dashboard.ProgressBar(i).Visible = False
        Case 5
            Dashboard.LabelTaskName(TaskCardIndex).Caption = "练习掌握度不佳的题目"
'            Navigator_Main.PictureIcon(TaskCardIndex).Picture = Navigator_Main.ImageList2.ListImages(10).Picture
            Dashboard.TaskActionBtn(TaskCardIndex).Tag = "5"
            Dashboard.ProgressBar(i).Visible = False
        Case 6
            Dashboard.LabelTaskName(TaskCardIndex).Caption = "练习高难题目"
'            Navigator_Main.PictureIcon(TaskCardIndex).Picture = Navigator_Main.ImageList2.ListImages(9).Picture
            Dashboard.TaskActionBtn(TaskCardIndex).Tag = "6"
            Dashboard.ProgressBar(i).Visible = False
        Case 7
            Dashboard.LabelTaskName(TaskCardIndex).Caption = "本题库的掌握度提升2%"
'            Navigator_Main.PictureIcon(TaskCardIndex).Picture = Navigator_Main.ImageList2.ListImages(10).Picture
            Dashboard.TaskActionBtn(TaskCardIndex).Visible = False
            Dashboard.ProgressBar(i).Visible = True
            Dashboard.ProgressBar(i).Max = 100
'            Navigator_Main.LabelFinish(i).Visible = False
        Case 8
            Dashboard.LabelTaskName(TaskCardIndex).Caption = "当日练习时长达20分钟"
'            Navigator_Main.PictureIcon(TaskCardIndex).Picture = Navigator_Main.ImageList2.ListImages(1).Picture
            Dashboard.TaskActionBtn(TaskCardIndex).Visible = False
            Dashboard.ProgressBar(i).Visible = True
            Dashboard.ProgressBar(i).Max = 20
'            Navigator_Main.LabelFinish(i).Visible = False
        Case 9
            Dashboard.LabelTaskName(TaskCardIndex).Caption = "复习高频的错题"
'            Navigator_Main.PictureIcon(TaskCardIndex).Picture = Navigator_Main.ImageList2.ListImages(2).Picture
            Dashboard.TaskActionBtn(TaskCardIndex).Tag = "9"
            Dashboard.ProgressBar(i).Visible = True
            Dashboard.ProgressBar(i).Max = 10
            Dashboard.TaskActionBtn(TaskCardIndex).Visible = True
'            Dashboard.LabelFinish(i).Visible = False
        Case 10
            Dashboard.LabelTaskName(TaskCardIndex).Caption = "以少于5个错误的成绩完成一次练习"
'            Dashboard.PictureIcon(TaskCardIndex).Picture = Navigator_Main.ImageList2.ListImages(5).Picture
            Dashboard.TaskActionBtn(TaskCardIndex).Tag = "10"
            
    End Select
Next i
'Navigator_Main.PictureIcon(0).Picture = Navigator_Main.ImageList2.ListImages(3).Picture

End If
End Function


Public Function checkLogValidity()
    Dim a As Object, fs As Object
    If Dir(App.Path & "\Archive\" & ChallengeLib & "\ChallengeLog") <> "" Then
        OpenTxt (App.Path & "\Archive\" & ChallengeLib & "\ChallengeLog")
        Dim block() As String
        block = Split(InputStr, vbCrLf)
        If block(0) <> Date Then
            Set fs = CreateObject("Scripting.FileSystemObject")
            Set a = fs.CreateTextFile(App.Path & "\Archive\" & ChallengeLib & "\ChallengeLog", True)
            a.WriteLine Date
            a.Close
            Set a = Nothing
            bb = RecordInitValue
        End If
    End If

If Dir(App.Path & "\Archive\" & ChallengeLib & "\topicHistory") <> "" Then
    If FileDateTime(App.Path & "\Archive\" & ChallengeLib & "\topicHistory") <> Date Then
        Set fs = CreateObject("Scripting.FileSystemObject")
        Set a = fs.CreateTextFile(App.Path & "\Archive\" & ChallengeLib & "\topicHistory", True)
        a.WriteLine ""
        a.Close
        Set a = Nothing
    End If

End If

End Function


Public Function challenge_searchEvents()
If Conditions = "" Then MsgBox ("错误：Condition参数为空。过程终止。"): Exit Function
DoEvents
'On Error Resume Next
ChallengeTemp.Show
DoEvents
Dim Condition() As String, Detail() As String, resultPooleach() As String, resultPoolinfo() As String
resultPool = ""
'预加载结果库内容，先加载全部内容的简要信息，再在之后的过程中逐个排除
Dim libPool As String
libPool = ""
Dim sb As String
Dim sf As String
Dim c As Integer
c = 0
sb = App.Path & "\Archive\"
sf = Dir(sb, vbDirectory)
'On Error Resume Next
Do While sf <> ""
If sf <> ".." And sf <> "." Then
    If (GetAttr(sb + sf) And vbDirectory) = vbDirectory Then
        libPool = libPool & sf & ":::"
        c = c + 1
    End If
End If
sf = Dir
Loop
Dim lib() As String
lib = Split(libPool, ":::")
For k = LBound(lib) To UBound(lib) - 1
    ChallengeTemp.File1.Path = App.Path & "\Archive\" & lib(k) & "\topics"
    For m = 1 To ChallengeTemp.File1.ListCount / 4
        If Dir(ChallengeTemp.File1.Path & "\" & m & "_topic") <> "" Then
            resultPool = resultPool & lib(k) & "::" & m & "}}"
        End If
    Next m
Next k
'预加载完毕
Debug.Print Conditions
Condition = Split(Conditions, "}}}")
Dim isFound As Boolean
For k = LBound(Condition) To UBound(Condition)
    resultPooleach = Split(resultPool, "}}")
    Detail = Split(Condition(k), ":::")
    If Condition(k) = "" Then Exit For
    If Detail(0) = "Kwords" Then
        For m = LBound(resultPooleach) To UBound(resultPooleach) - 1
            resultPoolinfo = Split(resultPooleach(m), "::")
            
            isFound = False
            If Detail(2) = "True" Then
                OpenTxt App.Path & "\Archive\" & resultPoolinfo(0) & "\topics\" & resultPoolinfo(1) & "_topic"
                If InStr(InputStr, Detail(1)) <> 0 Then isFound = True
            End If
            If Detail(3) = "True" Then
                OpenTxt App.Path & "\Archive\" & resultPoolinfo(0) & "\answers\" & resultPoolinfo(1) & "_answer"
                If InStr(InputStr, Detail(1)) <> 0 Then isFound = True
            End If
            If Detail(4) = "True" Then
                OpenTxt App.Path & "\Archive\" & resultPoolinfo(0) & "\resolve\" & resultPoolinfo(1) & "_resolve"
                If InStr(InputStr, Detail(1)) <> 0 Then isFound = True
            End If
            If isFound = False Then resultPool = Replace(resultPool, resultPooleach(m) & "}}", "")
        Next m
    ElseIf Detail(0) = "Lib" Then
        For m = LBound(resultPooleach) To UBound(resultPooleach) - 1
            resultPoolinfo = Split(resultPooleach(m), "::")
            isFound = False
            For p = 1 To UBound(Detail) - 1
                If Detail(p) = resultPoolinfo(0) Then isFound = True
            Next p
            If isFound = False Then resultPool = Replace(resultPool, resultPooleach(m) & "}}", "")
        Next m

    ElseIf Detail(0) = "Tags" Then
        For m = LBound(resultPooleach) To UBound(resultPooleach) - 1
            resultPoolinfo = Split(resultPooleach(m), "::")
            OpenTxt App.Path & "\Archive\" & resultPoolinfo(0) & "\topics\" & resultPoolinfo(1) & "_tags"
            isFound = False
            InputStr = Replace(InputStr, vbCrLf, "")
            Dim d() As String
            d = Split(Detail(1), ";")
            For t = LBound(d) To UBound(d) - 1
                If InStr(InputStr, d(t)) <> 0 Then isFound = True
            Next t
            If isFound = False Then resultPool = Replace(resultPool, resultPooleach(m) & "}}", "")
        Next m
    ElseIf Detail(0) = "Value" Then
        For m = LBound(resultPooleach) To UBound(resultPooleach) - 1
            resultPoolinfo = Split(resultPooleach(m), "::")
            OpenTxt App.Path & "\Archive\" & resultPoolinfo(0) & "\topics\" & resultPoolinfo(1) & "_stars"
            InputStr = Replace(InputStr, vbCrLf, "")
            isFound = False
            If Detail(2) = "bigger" Then
                If Val(InputStr) > Val(Detail(1)) Then isFound = True
            ElseIf Detail(2) = "smaller" Then
                If Val(InputStr) < Val(Detail(1)) Then isFound = True
            Else
                If Val(InputStr) = Val(Detail(1)) Then isFound = True
            End If
            If isFound = False Then resultPool = Replace(resultPool, resultPooleach(m) & "}}", "")
        Next m
    ElseIf Detail(0) = "Records" Then
        For m = LBound(resultPooleach) To UBound(resultPooleach) - 1
            resultPoolinfo = Split(resultPooleach(m), "::")
            isFound = False
            If Dir(App.Path & "\Archive\" & resultPoolinfo(0) & "\PracticeRecords", vbDirectory) = "" Then MkDir (App.Path & "\Archive\" & resultPoolinfo(0) & "\PracticeRecords")
            If Dir(App.Path & "\Archive\" & resultPoolinfo(0) & "\PracticeRecords\" & resultPoolinfo(1) & "_records") <> "" Then
                OpenTxt App.Path & "\Archive\" & resultPoolinfo(0) & "\PracticeRecords\" & resultPoolinfo(1) & "_records"
                InputStr = Replace(InputStr, vbCrLf, "")
                Dim practiceTime() As String, practiceInfo() As String
                practiceTime = Split(InputStr, "}}}")
                If Detail(1) = "TimePracticed" Then
                    If Detail(3) = "bigger" Then
                        If UBound(practiceTime) > Val(Detail(2)) Then isFound = True
                    ElseIf Detail(3) = smaller Then
                        If UBound(practiceTime) < Val(Detail(2)) Then isFound = True
                    Else
                        If UBound(practiceTime) = Val(Detail(2)) Then isFound = True
                    End If
                ElseIf Detail(1) = "TimeFault" Then
                    counter = 0
                    For U = LBound(practiceTime) To UBound(practiceTime) - 1
                        practiceInfo = Split(practiceTime(U), ":::")
                        If practiceInfo(2) = "false" Then counter = counter + 1
                    Next U
                    If Detail(3) = "bigger" Then
                        If counter > Val(Detail(2)) Then isFound = True
                    ElseIf Detail(3) = smaller Then
                        If counter < Val(Detail(2)) Then isFound = True
                    Else
                        If counter = Val(Detail(2)) Then isFound = True
                    End If
                ElseIf Detail(1) = "avrCorrect" Then
                    avr = 0
                    For U = LBound(practiceTime) To UBound(practiceTime) - 1
                        practiceInfo = Split(practiceTime(U), ":::")
                        avr = avr + Sin(practiceInfo(5))
                    Next U
                    avr = (avr / UBound(practiceTime)) * 100
                    If Detail(3) = "bigger" Then
                        If avr > Val(Detail(2)) Then isFound = True
                    ElseIf Detail(3) = "smaller" Then
                        If avr < Val(Detail(2)) Then isFound = True
                    Else
                        If avr = Val(Detail(2)) Then isFound = True
                    End If
                End If
            '
            Else
                If Detail(1) = "TimePracticed" And Detail(3) = "equals" And Val(Detail(2)) = 0 Then isFound = True
            '
            End If
                
            If isFound = False Then resultPool = Replace(resultPool, resultPooleach(m) & "}}", "")
        Next m
    ElseIf Detail(0) = "DTModified" Then
                    
                    For m = LBound(resultPooleach) To UBound(resultPooleach) - 1
                        resultPoolinfo = Split(resultPooleach(m), "::")
                        If Detail(3) = "lastEdited" Then
                            datecheck = FileDateTime(App.Path & "\Archive\" & resultPoolinfo(0) & "\topics\" & resultPoolinfo(1) & "_topic")
                            isFound = False
                            If Detail(2) = "earlier" Then
                                If DateDiff("d", Detail(1), datecheck) < 0 Then isFound = True
                            ElseIf Detail(2) = "after" Then
                                If DateDiff("d", Detail(1), datecheck) > 0 Then isFound = True
                            Else
                                If DateDiff("d", Detail(1), datecheck) = 0 Then isFound = True
                            End If
                        ElseIf Detail(3) = "lastPracticed" Then
                            isFound = False
                            If Dir(App.Path & "\Archive\" & resultPoolinfo(0) & "\PracticeRecords", vbDirectory) <> "" Then
                                If Dir(App.Path & "\Archive\" & resultPoolinfo(0) & "\PracticeRecords\" & resultPoolinfo(1) & "_records") <> "" Then
                                    OpenTxt App.Path & "\Archive\" & resultPoolinfo(0) & "\PracticeRecords\" & resultPoolinfo(1) & "_records"
                                    InputStr = Replace(InputStr, vbCrLf, "")
                                    Dim recordNum() As String, recordDetail() As String
                                    recordNum = Split(InputStr, "}}}")
                                    recordDetail = Split(recordNum(UBound(recordNum) - 1), ":::")
                                    'Debug.Print DateDiff("d", recordDetail(0), Detail(1))
                                    If Detail(2) = "today" Then
                                        If DateDiff("d", recordDetail(0), Detail(1)) = 0 Then isFound = True
                                    Else
                                        If DateDiff("d", recordDetail(0), Detail(1)) > 0 Then isFound = True
                                    End If
                                End If
                            End If
                        End If
                        If isFound = False Then resultPool = Replace(resultPool, resultPooleach(m) & "}}", "")
                    Next m
    End If


Next k
'debug
'SpecifiedSearchOccupied = True
'MDITopicSearchResult.Show
End Function


Public Function Challenge_Task_classicShuffle()
'    TemplateExamGenerateNotify.Show
'    TemplateExamGenerateNotify.Tag = "everydayChallengeEvent++" & ChallengeLib
Dim getTop As Long, allProcessed As Boolean, getScore As Double, getLoad As Single, availableCount As Integer
allProcessed = False
currentLoad = 0
availableCount = 0
If maxTopicNum = 0 Then maxTopicNum = 1000
getTop = GetFileCount(App.Path & "\Archive\" & ChallengeLib & "\topics")
Dim arr As Variant
arr = GetShuffledArray(1, Int(getTop / 4))
resultPool = ""
        Do Until currentLoad > maxTopicNum Or allProcessed
            allProcessed = True
            For i = LBound(arr) To UBound(arr)
                If currentLoad >= maxTopicNum Then Exit Do
                msg = calculatePriorityScore(ChallengeLib, arr(i), getScore, getLoad)
                If getScore <> 0 Then
                    resultPool = resultPool & ChallengeLib & "::" & arr(i) & "}}"
                    currentLoad = currentLoad + getLoad
                    availableCount = availableCount + 1
                End If
            Next i
            allProcessed = True
        Loop
        ab = GenerateExerciseFromTask(2, availableCount)
End Function

Public Function getFaultcount(libname As String) As Integer

Dim fso As Object, FolderPath As String
FolderPath = App.Path & "\Archive\" & libname & "\topics"
Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(FolderPath) Then
        MsgBox "无效题库", vbCritical
        getFaultcount = 0
        Exit Function
    End If
Dim targetFolder As String
Dim fileCount As Long
targetFolder = FolderPath
fileCount = GetFileCount(targetFolder)
If Dir(App.Path & "\Archive\" & libname & "\PracticeRecords", vbDirectory) = "" Then
    getFaultcount = 0
    Exit Function
Else
    msg = GetFileCount(App.Path & "\Archive\" & libname & "\PracticeRecords")
    If msg = 0 Then
       getFaultcount = 0
        Exit Function
    End If
End If
'初始化一张错题表，格式：题编号:::错误次数:::最新错误记录距今天数
Dim listOfErrors As String, trialCount As Integer, wrongCount As Integer, lastErrorInterval As Integer, lastErrorDate As Date
listOfErrors = ""
For i = 1 To Int(fileCount / 4)
    If Dir(App.Path & "\Archive\" & libname & "\topics\" & i & "_topic") <> "" Then
        If Dir(App.Path & "\Archive\" & libname & "\PracticeRecords\" & i & "_records") <> "" Then
            OpenTxt App.Path & "\Archive\" & libname & "\PracticeRecords\" & i & "_records"
            InputStr = Replace(InputStr, vbCrLf, "")
            Dim recordNum() As String, recordDetail() As String
            recordNum = Split(InputStr, "}}}")
            trialCount = UBound(recordNum)
            For j = LBound(recordNum) To UBound(recordNum)
                recordNum(j) = Replace(recordNum(j), vbCrLf, "")
                If Len(recordNum(j)) > 4 Then '有效记录，而非空格占位符
                    recordDetail = Split(recordNum(j), ":::")
                    If recordDetail(2) = "false" Then
                        getFaultcount = getFaultcount + 1
                        Exit For
                    End If
                End If
            Next j
        End If
    End If
Next i
Debug.Print getFaultcount



End Function



Public Function getListofRecentFaults(libname As String)
Dim fso As Object, FolderPath As String
FolderPath = App.Path & "\Archive\" & libname & "\topics"
Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(FolderPath) Then
        MsgBox "无效题库", vbCritical
        Exit Function
    End If
Dim targetFolder As String
Dim fileCount As Long
    ' 目标文件夹（示例：App.Path，可替换为任意路径）
targetFolder = FolderPath
    ' 1. 获取所有文件数量
fileCount = GetFileCount(targetFolder)
If Dir(App.Path & "\Archive\" & libname & "\PracticeRecords", vbDirectory) = "" Then
'    MsgBox "当前题库没有练习记录哦！点击“现在就练”，先练习一次吧！", vbInformation
    Exit Function
Else
    msg = GetFileCount(App.Path & "\Archive\" & libname & "\PracticeRecords")
    If msg = 0 Then
'        MsgBox "当前题库没有练习记录哦！点击“现在就练”，先练习一次吧！", vbInformation
        Exit Function
    End If
End If
'初始化一张错题表，格式：题编号:::错误次数:::最新错误记录距今天数
Dim listOfErrors As String, trialCount As Integer, wrongCount As Integer, lastErrorInterval As Integer, lastErrorDate As Date
listOfErrors = ""
For i = 1 To Int(fileCount / 4)
    If Dir(App.Path & "\Archive\" & libname & "\topics\" & i & "_topic") <> "" Then
        If Dir(App.Path & "\Archive\" & libname & "\PracticeRecords\" & i & "_records") <> "" Then
            OpenTxt App.Path & "\Archive\" & libname & "\PracticeRecords\" & i & "_records"
            InputStr = Replace(InputStr, vbCrLf, "")
            Dim recordNum() As String, recordDetail() As String
            recordNum = Split(InputStr, "}}}")
            trialCount = UBound(recordNum)
            For j = LBound(recordNum) To UBound(recordNum)
                recordNum(j) = Replace(recordNum(j), vbCrLf, "")
                If Len(recordNum(j)) > 4 Then '有效记录，而非空格占位符
                    recordDetail = Split(recordNum(j), ":::")
                    If recordDetail(2) = "false" Then
                        wrongCount = wrongCount + 1
                        lastErrorDate = recordDetail(0)
                        lastErrorInterval = DateDiff("D", lastErrorDate, Date)
                    End If
                End If
            Next j
            listOfErrors = listOfErrors & i & ":" & wrongCount & ":" & lastErrorInterval & "}"
            wrongCount = 0
            trialCount = 0
            lastErrorInterval = 0
            lastErrorDate = Date
        End If
    End If
Next i
Dim renewedListofError As String
renewedListofError = SortErrorList(listOfErrors, trainMode)
Debug.Print renewedListofError
getListofRecentFaults = renewedListofError


End Function

Public Function Challenge_Task_recentFaults(trainMode As Integer)

'get a list of recentFaults


'根据maxTopicNum，整合练习清单，发送至GenerateExerciseFromTask(3)
resultPool = ""
Dim renewedBlock() As String, renewedUnit() As String, availableCount As Integer
availableCount = 0
currentLoad = 0
renewedBlock = Split(getListofRecentFaults(ChallengeLib), "}")
If maxTopicNum = 0 Then maxTopicNum = 1000 '未读到最大负荷时的默认值
Dim getScore As Double, getLoad As Single, hasValidItem As Boolean

Select Case trainMode
    Case 1
        Do Until currentLoad > maxTopicNum
            hasValidItem = False
            For i = LBound(renewedBlock) To UBound(renewedBlock)
                If currentLoad > maxTopicNum Then Exit Do
                renewedUnit = Split(renewedBlock(i), ":")
                msg = calculatePriorityScore(ChallengeLib, Int(renewedUnit(0)), getScore, getLoad)
                If getScore <> 0 Then '为0表示今天练过了或者题目被禁用了。这么做防止多次申请练习结果出的是一样的题
                    resultPool = resultPool & ChallengeLib & "::" & renewedUnit(0) & "}}"
                    currentLoad = currentLoad + getLoad
                    availableCount = availableCount + 1
                    hasValidItem = True
                End If
            Next i
            If Not hasValidItem Then Exit Do
        Loop
        If availableCount = 0 Then
            msg = MsgBox("今天的错题已经练完了，恭喜你！明天再来吧！", vbInformation, "好厉害")
            Exit Function
        End If
        ab = GenerateExerciseFromTask(9, availableCount)
        
        
    Case 2
        Do Until currentLoad > maxTopicNum
            hasValidItem = False
            For i = UBound(renewedBlock) To LBound(renewedBlock) Step -1
                If currentLoad > maxTopicNum Then Exit Do
                renewedUnit = Split(renewedBlock(i), ":")
                msg = calculatePriorityScore(ChallengeLib, Int(renewedUnit(0)), getScore, getLoad)
                If getScore <> 0 Then
                    resultPool = resultPool & ChallengeLib & "::" & renewedUnit(0) & "}}"
                    currentLoad = currentLoad + getLoad
                    availableCount = availableCount + 1
                    hasValidItem = True
                End If
            Next i
            If Not hasValidItem Then Exit Do
        Loop
        If availableCount = 0 Then
            msg = MsgBox("今天的错题已经练完了，恭喜你！明天再来吧！", vbInformation, "好厉害")
            Exit Function
        End If
        ab = GenerateExerciseFromTask(3, availableCount)
End Select
Debug.Print resultPool

    '提纯完成，使用该resultPool创建练习，如果数量>25则限制题目数量为25，1天分2次完成，一天不超过50题

    '检查智能体推荐文件，该题推荐系数是否为0，若是，则代表今日已做过，删除。
    
    '结束检查
'    Dim shuffledItem() As String
'    shuffledItem = Split(resultPool, "}}")
'    Randomize timer ' 初始化随机数生成器
'    Dim temp As String
'    For i = UBound(shuffledItem) - 1 To LBound(shuffledItem) Step -1
'        j = Int((i - LBound(shuffledItem)) * Rnd + LBound(shuffledItem))
'        ' 交换元素
'        temp = shuffledItem(i)
'        shuffledItem(i) = shuffledItem(j)
'        shuffledItem(j) = temp
'    Next i
'    resultPool = ""
'    For i = LBound(shuffledItem) To UBound(shuffledItem)
'        resultPool = resultPool & shuffledItem(i) & "}}"
'    Next i
'    Debug.Print resultPool
'    Select Case itemCount
'        Case 0 To 25
'            ChallengeTopicMaximum = itemCount
'        Case Else
'            ChallengeTopicMaximum = 25
'    End Select

'    c = shuffleSetCount
    
End Function

Public Function Challenge_Task_unPracticed_7days()
Conditions = "Lib:::" & ChallengeLib & ":::}}}Records:::TimePracticed:::0:::equals}}}"
a = challenge_searchEvents

'处理resultPool
Dim block() As String, unit() As String
block = Split(resultPool, "}}")
Dim currentLoad As Integer, currentCount As Integer
currentLoad = 0
currentCount = 0
If maxTopicNum = 0 Then maxTopicNum = 1000
resultPool = ""
Dim getScore As Double, getLoad As Single
For i = LBound(block) To UBound(block)
    unit = Split(block(i), "::")
    msg = calculatePriorityScore(unit(0), Int(unit(1)), getScore, getLoad)
    currentLoad = currentLoad + getLoad
    currentCount = currentCount + 1
    resultPool = resultPool & unit(0) & "::" & unit(1) & "}}"
    If currentLoad >= maxTopicNum Then Exit For
Next i
If currentCount = 0 Then
    MsgBox ("恭喜！现在没有未练习的内容。" & vbCrLf & vbCrLf & "当你需要添加新内容时，可以通过“导入”添加。")
    If ChallengeTemp.Visible = True Then Unload ChallengeTemp
    Exit Function
End If

c = shuffleSetCount
ab = GenerateExerciseFromTask(4, currentCount)
End Function

Public Function Challenge_Task_accuracyBelowAverage()
'计算本题库平均正确率
ChallengeTemp.Show
DoEvents
ChallengeTemp.File1.Path = App.Path & "\Archive\" & ChallengeLib & "\topics"
Dim ratingLib() As Integer, n As Integer, low As Integer, midd As Integer, high As Integer, unPracticed As Integer, totalAvrAccuracy As Single, avrAccuracy As Single
totalAvrAccuracy = 0
If ChallengeTemp.File1.ListCount <> 0 Then
    Dim Rating As Single
    Rating = 0
    Dim itemCount As Integer
    itemCount = 0
    For i = 1 To ChallengeTemp.File1.ListCount / 4
        If Dir(App.Path & "\Archive\" & ChallengeLib & "\topics\" & i & "_topic") <> "" Then
        itemCount = itemCount + 1
        End If
    Next i
    ReDim ratingLib(1 To ChallengeTemp.File1.ListCount / 4)
    For k = 1 To ChallengeTemp.File1.ListCount / 4
        If Dir(App.Path & "\Archive\" & ChallengeLib & "\topics\" & k & "_topic") <> "" Then
            If Dir(App.Path & "\Archive\" & ChallengeLib & "\PracticeRecords\" & k & "_records") <> "" Then
                OpenTxt App.Path & "\Archive\" & ChallengeLib & "\PracticeRecords\" & k & "_records"
                InputStr = Replace(InputStr, vbCrLf, "")
                Dim recordNum() As String, recordDetail() As String
                recordNum = Split(InputStr, "}}}")
                Dim faultCount As Integer, totalPoint As Single
                faultCount = 0
                totalPoint = 0
                For kk = LBound(recordNum) To UBound(recordNum) - 1
                    recordDetail = Split(recordNum(kk), ":::")
                    If recordDetail(2) = "false" Then
                    
                        faultCount = faultCount + 1
                    End If
                    totalPoint = totalPoint + Val(recordDetail(5))
                Next kk
                totalAvrAccuracy = totalAvrAccuracy + Format(totalPoint / UBound(recordNum) * 100, "0.00")
            End If
'        OpenTxt App.Path & "\Archive\" & ChallengeLib & "\topics\" & k & "_stars"
'        InputStr = Replace(InputStr, vbCrLf, "")
'        star = Int(InputStr)
'        ratingLib(k) = star
'        If Dir(App.Path & "\Archive\" & ChallengeLib & "\PracticeRecords\" & k & "_records") <> "" Then
'            If ratingLib(k) < 41 Then
'                low = low + 1
'            ElseIf ratingLib(k) >= 41 And ratingLib(k) < 75 Then
'                midd = midd + 1
'            ElseIf ratingLib(k) >= 75 Then
'                high = high + 1
'            End If
'        Else
'            unPracticed = unPracticed + 1
'        End If
'        Rating = Rating + star
        End If
    Next k
    avrAccuracy = totalAvrAccuracy / itemCount
'遍历查找符合条件的题目
resultPool = ""
Dim currentCount As Integer
currentLoad = 0
currentCount = 0
If maxTopicNum = 0 Then maxTopicNum = 1000
    For k = 1 To ChallengeTemp.File1.ListCount / 4
        If Dir(App.Path & "\Archive\" & ChallengeLib & "\topics\" & k & "_topic") <> "" Then
            If Dir(App.Path & "\Archive\" & ChallengeLib & "\PracticeRecords\" & k & "_records") <> "" Then
                OpenTxt App.Path & "\Archive\" & ChallengeLib & "\PracticeRecords\" & k & "_records"
                InputStr = Replace(InputStr, vbCrLf, "")
                
                recordNum = Split(InputStr, "}}}")
                
                faultCount = 0
                totalPoint = 0
                For kk = LBound(recordNum) To UBound(recordNum) - 1
                    recordDetail = Split(recordNum(kk), ":::")
                    If recordDetail(2) = "false" Then
                        faultCount = faultCount + 1
                    End If
                    totalPoint = totalPoint + Val(recordDetail(5))
                Next kk
                If avrAccuracy > Format(totalPoint / UBound(recordNum) * 100, "0.00") Then
                    Dim getScore As Double, getLoad As Single
                    msg = calculatePriorityScore(ChallengeLib, k, getScore, getLoad)
                    If getScore <> 0 Then
                        currentCount = currentCount + 1
                        resultPool = resultPool & ChallengeLib & "::" & k & "}}"
                        currentLoad = currentLoad + getLoad
                        If currentLoad >= 1000 Then Exit For
                    End If
                End If
            End If
        End If
    Next k

c = shuffleSetCount
ab = GenerateExerciseFromTask(1, currentCount)

End If
'遍历题库每一题的平均正确率，若小于平均正确率，则加入resultPool

'执行generateExerciseFromTask函数

End Function

Public Function Challenge_Task_valueBelowAverage()
ChallengeTemp.Show
DoEvents
ChallengeTemp.File1.Path = App.Path & "\Archive\" & ChallengeLib & "\topics"
Dim ratingLib() As Integer, n As Integer, low As Integer, midd As Integer, high As Integer, unPracticed As Integer
If ChallengeTemp.File1.ListCount <> 0 Then
    Dim Rating As Single
    Rating = 0
    Dim itemCount As Integer
    itemCount = 0
    For i = 1 To ChallengeTemp.File1.ListCount / 4
        If Dir(App.Path & "\Archive\" & ChallengeLib & "\topics\" & i & "_topic") <> "" Then
        itemCount = itemCount + 1
        End If
    Next i
    ReDim ratingLib(1 To ChallengeTemp.File1.ListCount / 4)
    For k = 1 To ChallengeTemp.File1.ListCount / 4
        If Dir(App.Path & "\Archive\" & ChallengeLib & "\topics\" & k & "_topic") <> "" Then
        OpenTxt App.Path & "\Archive\" & ChallengeLib & "\topics\" & k & "_stars"
        InputStr = Replace(InputStr, vbCrLf, "")
        star = Int(InputStr)
        ratingLib(k) = star
        If Dir(App.Path & "\Archive\" & ChallengeLib & "\PracticeRecords\" & k & "_records") <> "" Then
            If ratingLib(k) < 41 Then
                low = low + 1
            ElseIf ratingLib(k) >= 41 And ratingLib(k) < 75 Then
                midd = midd + 1
            ElseIf ratingLib(k) >= 75 Then
                high = high + 1
            End If
        Else
            unPracticed = unPracticed + 1
        End If
        Rating = Rating + star
        End If
    Next k
    totalrating = Rating
    avrRating = totalrating / itemCount
'遍历查找符合条件的题目
resultPool = ""
Dim currentCount As Integer
currentLoad = 0
currentCount = 0
If maxTopicNum = 0 Then maxTopicNum = 1000
    For k = 1 To ChallengeTemp.File1.ListCount / 4
        If Dir(App.Path & "\Archive\" & ChallengeLib & "\topics\" & k & "_topic") <> "" Then
        OpenTxt App.Path & "\Archive\" & ChallengeLib & "\topics\" & k & "_stars"
        InputStr = Replace(InputStr, vbCrLf, "")
        star = Int(InputStr)
        Dim getScore As Double, getLoad As Single
        msg = calculatePriorityScore(ChallengeLib, Int(k), getScore, getLoad)
        currentLoad = currentLoad + getLoad
        If currentLoad >= maxTopicNum Then Exit For
        If star >= avrRating And getScore <> 0 Then resultPool = resultPool & ChallengeLib & "::" & k & "}}": currentCount = currentCount + 1
        End If
    Next k
c = shuffleSetCount
ab = GenerateExerciseFromTask(5, currentCount)
End If
End Function

Public Function Challenge_Task_highValue()
ChallengeTemp.Show
DoEvents
ChallengeTemp.File1.Path = App.Path & "\Archive\" & ChallengeLib & "\topics"
resultPool = ""
Dim currentLoad As Single, currentCount As Integer
currentCount = 0
currentLoad = 0
If maxTopicNum = 0 Then maxTopicNum = 1000
Dim ratingLib() As Integer, n As Integer, low As Integer, midd As Integer, high As Integer, unPracticed As Integer
If ChallengeTemp.File1.ListCount <> 0 Then
    Dim Rating As Single
    Rating = 0
    Dim itemCount As Integer
    itemCount = 0
'    For i = 1 To ChallengeTemp.File1.ListCount / 4
'        If Dir(App.Path & "\Archive\" & ChallengeLib & "\topics\" & i & "_topic") <> "" Then
'        itemCount = itemCount + 1
'        End If
'    Next i
    ReDim ratingLib(1 To ChallengeTemp.File1.ListCount / 4)
    For k = 1 To ChallengeTemp.File1.ListCount / 4
        If Dir(App.Path & "\Archive\" & ChallengeLib & "\topics\" & k & "_topic") <> "" Then
        OpenTxt App.Path & "\Archive\" & ChallengeLib & "\topics\" & k & "_stars"
        InputStr = Replace(InputStr, vbCrLf, "")
        star = Int(InputStr)
        ratingLib(k) = star
        If Dir(App.Path & "\Archive\" & ChallengeLib & "\PracticeRecords\" & k & "_records") <> "" Then
            Dim getScore As Double, getLoad As Single
            msg = calculatePriorityScore(ChallengeLib, Int(k), getScore, getLoad)
            If ratingLib(k) >= 75 And getScore <> 0 Then
                resultPool = resultPool & ChallengeLib & "::" & k & "}}"
                currentLoad = currentLoad + getLoad
                currentCount = currentCount + 1
                If currentLoad > maxTopicNum Then Exit For
            End If
        Else
            unPracticed = unPracticed + 1
        End If
        Rating = Rating + star
        End If
    Next k
'    totalRating = Rating
'    avrRating = totalRating / itemCount
'遍历查找符合条件的题目
If currentCount = 0 Then inf = MsgBox("太棒了！现在没有可供练习的难题。" & vbCrLf & vbCrLf & "试试“现在就练”，随便练点，说不定就有难题了呢...", vbInformation, "太棒了"): Unload ChallengeTemp: Exit Function
c = shuffleSetCount
ab = GenerateExerciseFromTask(6, currentCount)
End If

End Function
