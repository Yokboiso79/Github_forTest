
Module Start

    Sub Main()

        Do

            Dim LicProcess As Process()
            Dim AQLicProcess As Process()
            'Dim PDMUnUsedProcess As Process()

            LicProcess = Process.GetProcessesByName("LICMonitoringService")
            AQLicProcess = Process.GetProcessesByName("ABAQUS_USED_RECORD")
            'PDMUnUsedProcess = Process.GetProcessesByName("PDM_LIC_USER_Filter")

            If LicProcess.Length = 0 Then
                'SetApp("C:\LicenseMonitoring\LICMonitoringService.exe")
                Call Shell("C:\LicenseMonitoring\LICMonitoringService.exe", AppWinStyle.NormalFocus)
            End If
            If AQLicProcess.Length = 0 Then
                Call Shell("C:\ABAQUS Monitoring\ABAQUS_USED_RECORD.exe", AppWinStyle.NormalFocus)
            End If
            'If PDMUnUsedProcess.Length = 0 Then
            '    Call Shell("C:\PDMUnusedCheck\PDM_LIC_USER_Filter.exe", AppWinStyle.NormalFocus)
            'End If
            System.Threading.Thread.Sleep(300000)
        Loop

    End Sub

    Private Function SetApp(ByVal strFileName As String, Optional ByVal strArgument As String = Nothing) As Boolean

        Try
            Dim GetProcess As New Process
            GetProcess.StartInfo.FileName = strFileName
            GetProcess.StartInfo.Arguments = strArgument
            GetProcess.StartInfo.RedirectStandardInput = True
            GetProcess.StartInfo.RedirectStandardOutput = True
            GetProcess.StartInfo.UseShellExecute = False
            GetProcess.StartInfo.WindowStyle = ProcessWindowStyle.Normal
            GetProcess.Start()

            GetProcess.Dispose()

            Return True

        Catch ex As Exception
            Return False
        End Try

    End Function

End Module
