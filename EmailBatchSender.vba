Sub SendBulkEmails_Complete()
    ' 1. 声明变量
    Dim xlApp As Object
    Dim xlWB As Object
    Dim xlSheet As Object
    Dim i As Integer
    Dim LastRow As Long
    Dim MailItem As Outlook.MailItem
    Dim FilePath As String
    
    ' --- 【核心配置区：请修改路径】 ---
    ' 提示：右键点击 Excel 文件 -> 属性 -> 安全 -> 复制“对象名称”获取准确路径
    FilePath = "C:\Users\你的用户名\Desktop\MailList.xlsx"
    
    ' 2. 启动/连接 Excel
    On Error Resume Next
    Set xlApp = GetObject(, "Excel.Application")
    If xlApp Is Nothing Then
        Set xlApp = CreateObject("Excel.Application")
    End If
    On Error GoTo 0
    
    ' 3. 打开工作簿
    On Error GoTo ErrorHandler
    Set xlWB = xlApp.Workbooks.Open(FilePath, ReadOnly:=True)
    Set xlSheet = xlWB.Sheets(1)
    
    ' 计算最后一行
    LastRow = xlSheet.Cells(xlSheet.Rows.Count, 1).End(-4162).Row ' -4162 代表 xlUp
    
    ' 4. 循环处理
    For i = 2 To LastRow
        ' 提取 Excel 数据到局部变量
        Dim curName As String, curEmail As String, curSub As String
        curName = Trim(xlSheet.Cells(i, 1).Value)
        curEmail = Trim(xlSheet.Cells(i, 2).Value)
        curSub = Trim(xlSheet.Cells(i, 3).Value)
        
        ' 跳过邮箱为空的行
        If curEmail <> "" Then
            Set MailItem = Application.CreateItem(olMailItem)
            
            With MailItem
                .To = curEmail
                
                ' 设置主题 (如果 Excel 里没写，则使用默认主题)
                If curSub <> "" Then
                    .Subject = curSub
                Else
                    .Subject = "商务邮件问候"
                End If
                
                ' 设置正文 (HTML 格式，支持换行和加粗)
                .HTMLBody = "<html><body style='font-family:微软雅黑; font-size:14px;'>" & _
                            "<p>亲爱的 <b>" & curName & "</b>：</p>" & _
                            "<p>您好！这是一封根据您的需求自动生成的商务邮件。</p>" & _
                            "<p>附件信息已备好，请查收并审阅。如有任何疑问，欢迎随时联系我。</p>" & _
                            "<p><br>顺颂商祺！</p>" & _
                            "</body></html>"
                
                ' --- 测试阶段建议用 .Display（弹出预览） ---
                ' --- 正式群发请改为 .Send ---
                .Display 
                
            End With
            Set MailItem = Nothing
        End If
    Next i
    
    ' 5. 清理现场
    xlWB.Close SaveChanges:=False
    xlApp.Quit
    Set xlSheet = Nothing: Set xlWB = Nothing: Set xlApp = Nothing
    
    MsgBox "邮件批量处理已完成！共计: " & (LastRow - 1) & " 封。", vbInformation
    Exit Sub

ErrorHandler:
    MsgBox "发生意外错误！" & vbCrLf & "错误号: " & Err.Number & vbCrLf & "原因: " & Err.Description
End Sub
