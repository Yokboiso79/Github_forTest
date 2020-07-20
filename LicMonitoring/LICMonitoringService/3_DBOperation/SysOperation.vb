Imports System.IO
Imports System.Data.SqlClient


Public Class SysOperation

    'Public Shared InsertTime As Boolean

    Public Function CATIALicInfo_NEW(ByVal arrContents As ArrayList, ByVal HostName As String, ServerInfo As String) As ArrayList

        Try
            '현재 시간
            Dim StartTime As Date = My.Computer.Clock.LocalTime
            Dim CMDprocess As New Process
            Dim StartInfo As New System.Diagnostics.ProcessStartInfo
            StartInfo.FileName = "i4blt"
            StartInfo.Arguments = " -s -n " & HostName
            StartInfo.RedirectStandardInput = True
            StartInfo.RedirectStandardOutput = True
            StartInfo.UseShellExecute = False
            CMDprocess.StartInfo = StartInfo
            CMDprocess.Start()
            Dim SR As System.IO.StreamReader = CMDprocess.StandardOutput
            Dim SW As System.IO.StreamWriter = CMDprocess.StandardInput

            Do
                Dim GetLine As String = SR.ReadLine
                arrContents.Add(GetLine)
            Loop Until SR.EndOfStream = True

            SW.Close()
            SR.Close()

            '완료 시간
            Dim EndTime As Date = My.Computer.Clock.LocalTime
            '시간 비교 테스트
            Dim DifferenceInSeconds As Long = DateDiff(DateInterval.Second, StartTime, EndTime)
            'Dim SpanFromSeconds As TimeSpan = New TimeSpan(0, 0, DifferenceInSeconds)
            If DifferenceInSeconds >= 300 Then
                Dim strCategory As String = "CATIA 라이선스 서버 장애 발생 - " & HostName
                SendMail(strCategory, ServerInfo & vbNewLine & "5분간 라이선스 정보를 받아오지 못하고 있습니다. 서버 확인 바랍니다.")
            End If

            Return arrContents

        Catch ex As Exception

            Return arrContents

        End Try


    End Function


    Public Function DSLSLicInfo_NEW(ByVal arrContents As ArrayList, ByVal ServerIP As String)

        Try
            '현재 시간
            Dim StartTime As Date = My.Computer.Clock.LocalTime
            Dim CMDprocess As New Process
            Dim StartInfo As New System.Diagnostics.ProcessStartInfo
            Dim strArg As String = "-admin -run " & """c " & ServerIP & " 4084;getLicenseUsage -all"""
            'Dim strArg As String = "-admin -run " & """c " & ServerIP & " 4084;getLicenseUsage"""

            Debug.Print(strArg)
            StartInfo.FileName = "C:\Program Files\Dassault Systemes\DS License Server\win_b64\code\bin\dslicsrv.exe"
            StartInfo.Arguments = strArg
            StartInfo.RedirectStandardInput = True
            StartInfo.RedirectStandardOutput = True
            StartInfo.UseShellExecute = False
            CMDprocess.StartInfo = StartInfo
            CMDprocess.Start()
            Dim SR As System.IO.StreamReader = CMDprocess.StandardOutput
            Dim SW As System.IO.StreamWriter = CMDprocess.StandardInput

            Do
                Dim GetLine As String = SR.ReadLine
                arrContents.Add(GetLine)
                Debug.Print(GetLine)
            Loop Until SR.EndOfStream = True

            SW.Close()
            SR.Close()

            '완료 시간
            Dim EndTime As Date = My.Computer.Clock.LocalTime
            '시간 비교 테스트
            Dim DifferenceInSeconds As Long = DateDiff(DateInterval.Second, StartTime, EndTime)
            'Dim SpanFromSeconds As TimeSpan = New TimeSpan(0, 0, DifferenceInSeconds)
            'If DifferenceInSeconds >= 300 Then
            '    Dim strCategory As String = "CATIA 라이선스 서버 장애 발생 - " & HostName
            '    SendMail(strCategory, ServerInfo & vbNewLine & "5분간 라이선스 정보를 받아오지 못하고 있습니다. 서버 확인 바랍니다.")
            'End If



        Catch ex As Exception


        End Try

        Return arrContents

    End Function
    Public Function CATIALicInfo_DSLS(ByVal arrContents As ArrayList, ByVal HostName As String, ServerInfo As String) As ArrayList

        Try
            '현재 시간
            Dim StartTime As Date = My.Computer.Clock.LocalTime
            Dim CMDprocess As New Process
            Dim StartInfo As New System.Diagnostics.ProcessStartInfo
            Dim strArg As String = "-admin -run " & """c " & HostName & " 4084;19450815;getLicenseUsage -all"""
            StartInfo.FileName = "C:\Program Files\Dassault Systemes\DS License Server\win_b64\code\bin\dslicsrv.exe"
            StartInfo.Arguments = strArg
            StartInfo.RedirectStandardInput = True
            StartInfo.RedirectStandardOutput = True
            StartInfo.UseShellExecute = False
            CMDprocess.StartInfo = StartInfo
            CMDprocess.Start()
            Dim SR As System.IO.StreamReader = CMDprocess.StandardOutput
            Dim SW As System.IO.StreamWriter = CMDprocess.StandardInput

            Do
                Dim GetLine As String = SR.ReadLine
                arrContents.Add(GetLine)
            Loop Until SR.EndOfStream = True

            SW.Close()
            SR.Close()

            '완료 시간
            Dim EndTime As Date = My.Computer.Clock.LocalTime
            '시간 비교 테스트
            Dim DifferenceInSeconds As Long = DateDiff(DateInterval.Second, StartTime, EndTime)
            'Dim SpanFromSeconds As TimeSpan = New TimeSpan(0, 0, DifferenceInSeconds)
            If DifferenceInSeconds >= 300 Then
                Dim strCategory As String = "CATIA 라이선스 서버 장애 발생 - " & HostName
                SendMail(strCategory, ServerInfo & vbNewLine & "5분간 라이선스 정보를 받아오지 못하고 있습니다. 서버 확인 바랍니다.")
            End If

            Return arrContents

        Catch ex As Exception

            Return arrContents

        End Try


    End Function

    Public Function CATIALicInfo(ByVal arrContents As ArrayList, ByVal AddONVer As String) As ArrayList

        Try

            Dim CMDprocess As New Process
            Dim StartInfo As New System.Diagnostics.ProcessStartInfo
            StartInfo.FileName = "i4blt"
            StartInfo.Arguments = " -s"
            StartInfo.RedirectStandardInput = True
            StartInfo.RedirectStandardOutput = True
            StartInfo.UseShellExecute = False
            CMDprocess.StartInfo = StartInfo
            CMDprocess.Start()
            Dim SR As System.IO.StreamReader = CMDprocess.StandardOutput
            Dim SW As System.IO.StreamWriter = CMDprocess.StandardInput

            Do
                Dim GetLine As String = SR.ReadLine
                If GetLine.Contains("Product Version:  R1") Then
                    GetLine = GetLine.Replace("R1", AddONVer & "_R1")
                End If
                arrContents.Add(GetLine)
            Loop Until SR.EndOfStream = True

            SW.Close()
            SR.Close()

            Return arrContents

        Catch ex As Exception

            Return arrContents

        End Try


    End Function

    Public Function NXLicInfo(ByVal arrContents As ArrayList, ByVal lmUtilPath As String, ByVal strArg As String, Optional ByVal AddONVer As String = Nothing) As ArrayList

        Try
            Dim CMDprocess As New Process
            Dim StartInfo As New System.Diagnostics.ProcessStartInfo
            'StartInfo.FileName = "C:\DynavistaLCS\lmutil.exe"
            StartInfo.FileName = lmUtilPath
            StartInfo.Arguments = strArg
            StartInfo.RedirectStandardInput = True
            StartInfo.RedirectStandardOutput = True
            StartInfo.UseShellExecute = False
            CMDprocess.StartInfo = StartInfo
            CMDprocess.Start()
            Dim SR As System.IO.StreamReader = CMDprocess.StandardOutput
            Dim SW As System.IO.StreamWriter = CMDprocess.StandardInput

            Do
                Dim GetLine As String = SR.ReadLine
                If GetLine.Contains("NX13100N") And GetLine.Contains("v30.0") Then
                    GetLine = GetLine.Replace("v30.0", "v30.0" & AddONVer)
                End If
                arrContents.Add(GetLine)
            Loop Until SR.EndOfStream = True

            SW.Close()
            SR.Close()

            Return arrContents

        Catch ex As Exception

            Return arrContents

        End Try


    End Function

    Public Function NXLicInfo_NEW(ByVal arrContents As ArrayList, ByVal lmUtilPath As String, ByVal strArg As String, ServerInfo As String) As ArrayList

        Try
            '현재 시간
            Dim StartTime As Date = My.Computer.Clock.LocalTime

            Dim CMDprocess As New Process
            Dim StartInfo As New System.Diagnostics.ProcessStartInfo
            StartInfo.FileName = lmUtilPath
            StartInfo.Arguments = strArg
            StartInfo.RedirectStandardInput = True
            StartInfo.RedirectStandardOutput = True
            StartInfo.UseShellExecute = False
            CMDprocess.StartInfo = StartInfo
            CMDprocess.Start()
            Dim SR As System.IO.StreamReader = CMDprocess.StandardOutput
            Dim SW As System.IO.StreamWriter = CMDprocess.StandardInput

            Do
                Dim GetLine As String = SR.ReadLine
                arrContents.Add(GetLine)
            Loop Until SR.EndOfStream = True

            SW.Close()
            SR.Close()


            '완료 시간
            Dim EndTime As Date = My.Computer.Clock.LocalTime
            '시간 비교 테스트
            Dim DifferenceInSeconds As Long = DateDiff(DateInterval.Second, StartTime, EndTime)
            'Dim SpanFromSeconds As TimeSpan = New TimeSpan(0, 0, DifferenceInSeconds)
            If DifferenceInSeconds >= 300 Then
                Dim strCategory As String = "FlexLM 라이선스 서버 장애 발생"
                SendMail(strCategory, ServerInfo & vbNewLine & "5분간 라이선스 정보를 받아오지 못하고 있습니다. 서버 확인 바랍니다.")
            End If

            Return arrContents

        Catch ex As Exception

            Return arrContents

        End Try


    End Function

    Public Function ReadingFile() As ArrayList

        Dim SettingFileName As String = "Settings.ini"
        Dim SettingFilePath As String = My.Application.Info.DirectoryPath + "\" + SettingFileName

        Dim FileInfoReader As StreamReader
        Dim SettingInfo As New ArrayList
        If My.Computer.FileSystem.FileExists(SettingFilePath) = True Then
            FileInfoReader = My.Computer.FileSystem.OpenTextFileReader(SettingFilePath, System.Text.Encoding.GetEncoding("euc-kr"))
            Do
                Dim ReadLine As String = FileInfoReader.ReadLine
                SettingInfo.Add(ReadLine)
            Loop While FileInfoReader.EndOfStream = False
            FileInfoReader.Close()
        Else
            'MessageBox.Show("환경 파일이 없습니다." + vbNewLine + "관리자에게 문의 바랍니다.", "SL License Monitoring System")
            End
        End If

        Return SettingInfo

    End Function

    Public Function ReadingFile_TST(ByVal SettingInfo As ArrayList, ByVal FilePath As String) As ArrayList

        Dim FileInfoReader As StreamReader
        If My.Computer.FileSystem.FileExists(FilePath) = True Then
            FileInfoReader = My.Computer.FileSystem.OpenTextFileReader(FilePath, System.Text.Encoding.GetEncoding("euc-kr"))
            Do
                Dim ReadLine As String = FileInfoReader.ReadLine
                SettingInfo.Add(ReadLine)
            Loop While FileInfoReader.EndOfStream = False
            FileInfoReader.Close()
        Else
            'MessageBox.Show("환경 파일이 없습니다." + vbNewLine + "관리자에게 문의 바랍니다.", "SL License Monitoring System")
            End
        End If

        Return SettingInfo

    End Function

    Public Function GetValue(ByVal strArr() As String) As ArrayList

        Dim arrValue As New ArrayList
        For index As Integer = 0 To strArr.Count - 1
            If strArr(index) <> "" Then
                arrValue.Add(strArr(index))
            End If
        Next

        Return arrValue

    End Function

    Public Function SortDateTime(ByVal CurrentTime As Date) As String

        Dim SQLDataTime As String
        SQLDataTime = CStr(CurrentTime.Year) + "-" + CStr(CurrentTime.Month) + "-" + CStr(CurrentTime.Day)
        SQLDataTime = SQLDataTime + " " + CStr(CurrentTime.Hour) + ":" + CStr(CurrentTime.Minute) + ":" + CStr(CurrentTime.Second)
        Debug.Print(SQLDataTime)

        Return SQLDataTime

    End Function

    Public Function SendMail(ByVal strCategory As String, ByVal Constents As String) As Boolean

        Dim strSubject As String = "License Mornitoring System : " & strCategory

        Dim Mail As New System.Net.Mail.MailAddress("systemmanager@slworld.com")
        Try
            Using SendMessage As New System.Net.Mail.MailMessage("systemmanager@slworld.com", "jaehwannoh@slworld.com", strSubject, Constents)
                'SendMessage.To.Add("jaehwannoh@slworld.com")
                'SendMessage.To.Add("choyeonghan@slworld.com")
                'SendMessage.To.Add("shkeum@slworld.com")
                'SendMessage.To.Add("heyheo@slworld.com")
                Dim MailClient As New System.Net.Mail.SmtpClient("210.105.188.9")
                MailClient.DeliveryMethod = System.Net.Mail.SmtpDeliveryMethod.Network
                MailClient.EnableSsl = False
                MailClient.UseDefaultCredentials = False
                MailClient.Credentials = New System.Net.NetworkCredential(Mail.User, "SLSystem")
                MailClient.Send(SendMessage)
            End Using
            Return True

        Catch ex As Exception
            Return False
        End Try

    End Function
    Public Function SendMail_forError(ByVal strCategory As String, ByVal Constents As String) As Boolean

        Dim strSubject As String = "License Mornitoring System : " & strCategory

        Dim Mail As New System.Net.Mail.MailAddress("systemmanager@slworld.com")
        Try
            Using SendMessage As New System.Net.Mail.MailMessage("systemmanager@slworld.com", "jaehwannoh@slworld.com", strSubject, Constents)

                Dim MailClient As New System.Net.Mail.SmtpClient("210.105.188.9")
                MailClient.DeliveryMethod = System.Net.Mail.SmtpDeliveryMethod.Network
                MailClient.EnableSsl = False
                MailClient.UseDefaultCredentials = False
                MailClient.Credentials = New System.Net.NetworkCredential(Mail.User, "SLSystem")
                MailClient.Send(SendMessage)
            End Using
            Return True

        Catch ex As Exception
            Return False
        End Try

    End Function

End Class
