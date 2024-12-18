Imports System.IO
Imports System.Net.Http
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Configuration
Imports Newtonsoft.Json

Public Class Form1
    Private watcher As FileSystemWatcher
    Private httpClient As HttpClient = New HttpClient()

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' 설정값 불러오기
        TextBox1.Text = ConfigurationManager.AppSettings("FolderPath")
        TextBox2.Text = ConfigurationManager.AppSettings("WebhookUrl")
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim folderPath As String = TextBox1.Text
        Dim webhookUrl As String = TextBox2.Text

        If String.IsNullOrWhiteSpace(folderPath) OrElse String.IsNullOrWhiteSpace(webhookUrl) Then
            MessageBox.Show("폴더 경로와 Webhook URL을 입력해주세요.")
            Return
        End If

        If Not Directory.Exists(folderPath) Then
            MessageBox.Show("지정된 폴더가 존재하지 않습니다.")
            Return
        End If

        ' 설정값 저장
        SaveSettings("FolderPath", folderPath)
        SaveSettings("WebhookUrl", webhookUrl)

        ' 파일 감시 설정
        watcher = New FileSystemWatcher()
        watcher.Path = folderPath
        watcher.Filter = "*.png"
        watcher.IncludeSubdirectories = True
        watcher.NotifyFilter = NotifyFilters.FileName Or NotifyFilters.CreationTime
        AddHandler watcher.Created, AddressOf OnNewImageCreated
        watcher.EnableRaisingEvents = True

        InvokeIfRequired(Sub() ListBox1.Items.Add("감시를 시작했습니다: " & folderPath))
    End Sub

    Private Sub OnNewImageCreated(sender As Object, e As FileSystemEventArgs)
        ' 이미지가 추가되었을 때 로그 파일을 확인하고 월드 이름 추출
        Try
            Dim baseLogFolder As String = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData).Replace("Local", "LocalLow"), "VRChat", "VRChat")

            If Not Directory.Exists(baseLogFolder) Then
                InvokeIfRequired(Sub() ListBox1.Items.Add("로그 폴더를 찾을 수 없습니다: " & baseLogFolder))
                Return
            End If

            Dim logFiles = Directory.GetFiles(baseLogFolder, "output_log_*.txt", SearchOption.AllDirectories)

            If logFiles.Length = 0 Then
                InvokeIfRequired(Sub() ListBox1.Items.Add("로그 파일을 찾을 수 없습니다."))
                Return
            End If

            Dim latestLogFile = logFiles.OrderByDescending(Function(f) New FileInfo(f).LastWriteTime).FirstOrDefault()

            If latestLogFile IsNot Nothing Then
                Dim worldName As String = ExtractWorldNameFromLog(latestLogFile)
                If Not String.IsNullOrEmpty(worldName) Then
                    SendToDiscord(TextBox2.Text, worldName, e.FullPath)
                End If
            End If
        Catch ex As Exception
            InvokeIfRequired(Sub() ListBox1.Items.Add("오류: " & ex.Message))
        End Try
    End Sub

    Private Function ExtractWorldNameFromLog(logFilePath As String) As String
        Try
            ' 읽기 전용으로 파일 열기
            Using fs As New FileStream(logFilePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)
                Using reader As New StreamReader(fs)
                    Dim lines = reader.ReadToEnd().Split(Environment.NewLine)
                    For Each line In lines.Reverse()
                        If line.Contains("Entering Room:") Then
                            Dim match = Regex.Match(line, "Entering Room: (.+)")
                            If match.Success Then
                                Dim worldName = match.Groups(1).Value
                                InvokeIfRequired(Sub() ListBox1.Items.Add("월드 이름 감지: " & worldName))
                                Return worldName
                            End If
                        End If
                    Next
                End Using
            End Using
        Catch ex As Exception
            InvokeIfRequired(Sub() ListBox1.Items.Add("로그 분석 중 오류: " & ex.Message))
        End Try

        Return String.Empty
    End Function

    Private Async Sub SendToDiscord(webhookUrl As String, worldName As String, imagePath As String)
        Try
            Dim captureTime As String = File.GetCreationTime(imagePath).ToString("yyyy-MM-dd HH:mm:ss")

            ' JSON 형식으로 전송 내용 구성
            Dim jsonPayload As New With {
                Key .content = "새로운 이미지가 업로드되었습니다.",
                Key .embeds = New Object() {
                    New With {
                        Key .title = "이미지 정보",
                        Key .fields = New Object() {
                            New With {Key .name = "월드 이름", Key .value = worldName, Key .inline = True},
                            New With {Key .name = "촬영 일시", Key .value = captureTime, Key .inline = True}
                        }
                    }
                }
            }

            Dim jsonString As String = JsonConvert.SerializeObject(jsonPayload)
            Dim content = New StringContent(jsonString, Encoding.UTF8, "application/json")

            ' 이미지를 첨부
            Dim boundary As String = "----WebKitFormBoundary" & DateTime.Now.Ticks.ToString("x")
            Dim multipartContent = New MultipartFormDataContent(boundary)
            multipartContent.Add(content, "payload_json")
            multipartContent.Add(New ByteArrayContent(File.ReadAllBytes(imagePath)), "file", Path.GetFileName(imagePath))

            Dim response = Await httpClient.PostAsync(webhookUrl, multipartContent)
            If response.IsSuccessStatusCode Then
                InvokeIfRequired(Sub() ListBox1.Items.Add("Discord로 전송 성공: " & worldName))
            Else
                InvokeIfRequired(Sub() ListBox1.Items.Add("Discord 전송 실패: " & response.StatusCode))
            End If
        Catch ex As Exception
            InvokeIfRequired(Sub() ListBox1.Items.Add("Discord 전송 오류: " & ex.Message))
        End Try
    End Sub

    Private Sub InvokeIfRequired(action As Action)
        If Me.InvokeRequired Then
            Me.Invoke(action)
        Else
            action()
        End If
    End Sub

    Private Sub SaveSettings(key As String, value As String)
        Dim config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None)
        If config.AppSettings.Settings(key) Is Nothing Then
            config.AppSettings.Settings.Add(key, value)
        Else
            config.AppSettings.Settings(key).Value = value
        End If
        config.Save(ConfigurationSaveMode.Modified)
        ConfigurationManager.RefreshSection("appSettings")
    End Sub

    Private Sub Label1_Click(sender As Object, e As EventArgs) Handles Label1.Click

    End Sub
End Class
