Imports System.Threading
Imports System.IO
Imports System.Data.SqlClient

Module StartMonitoring

    Dim oSysOp As New SysOperation


    Dim IntervalTime As Integer = 10000
    Dim _CATIA As New CATIALicSrvCheck
    Dim _NX As New NXLicSrvCheck
    Dim _DB As New DBConnect
    Dim InsertTime As Boolean
    Dim InsertDate As Date = My.Computer.Clock.LocalTime
    Dim SearchTime As System.Threading.Thread

    Sub Main()

        'On Error Resume Next

        Console.WriteLine("========== 라이선스 모니터링을 시작합니다                             ==========")
        Console.WriteLine("========== Version : 3.3 , SL Corporation by Dongwan Hong             ==========")
        Console.WriteLine("========== 20150917 - ALM 라이선스 모니터링 추가                      ==========")
        Console.WriteLine("========== 20150922 - LIC_INFO 정보 기록 매일 12시                    ==========")
        Console.WriteLine("========== 20151021 - 서버에 없는 라이선스 타입 에러 수정             ==========")
        Console.WriteLine("========== 20151105 - 메일링, 사용자,조직 동기화 추가                 ==========")
        Console.WriteLine("==========          - 라이선스 서버별 조회로 변경                     ==========")
        Console.WriteLine("========== 20151116 - 사용자 확인 로직 변경                           ==========")
        Console.WriteLine("========== 20151121 - GMS4050 라이선스 버전 변경                      ==========")
        Console.WriteLine("========== 20151123 - 알수 없는 사용자 추가                           ==========")
        Console.WriteLine("==========            (오프라인 유저, 중국 미인증 장비)               ==========")
        Console.WriteLine("========== 20151129 - CATIA 미국 라이선스 추가                        ==========")
        Console.WriteLine("========== 20151210 - NX 미국 라이선스 추가                           ==========")
        Console.WriteLine("========== 20151215 - DEPT_ID 추가                                    ==========")
        Console.WriteLine("========== 20160106 - 알수없는 사용자 IPPlus에서 가져오기             ==========")
        Console.WriteLine("========== 20160314 - 금형 IP 변경(203.250.10.25 -> 210.105.97.18     ==========")
        Console.WriteLine("========== 20160601 - 라이선스 서버 변경                              ==========")
        Console.WriteLine("========== 20160623 - 3DCS 서버 변경                                  ==========")
        Console.WriteLine("========== 20160725 - Change ALLEGRO Version                          ==========")
        Console.WriteLine("========== 20160726 - 라이선스 정보 5분간 못받아 올경우 전체 메일 발송==========")
        Console.WriteLine("========== 20160831 - KDS 라이선스 서버 제외                          ==========")
        Console.WriteLine("========== 20161007 - MATLAB/MOLDFLOW/CODESCROLL 모니터링 추가        ==========")
        Console.WriteLine("========== 20161021 - FLUENT/HYPERWORKS 모니터링 추가                 ==========")
        Console.WriteLine("========== 20161025 - ANSYS 모니터링 추가                             ==========")
        Console.WriteLine("========== 20161228 - INTERFACE 추가                                  ==========")
        Console.WriteLine("========== 20170514 - LUM -> DSLS 모니터링으로 변경                   ==========")
        Console.WriteLine("========== 20180103 - CATIA 중경 라이선스 추가                        ==========")
        Console.WriteLine("========== 20180409 - CATIA V6 금형 라이선스 서버 변경                ==========")
        Console.WriteLine("========== 20180629 - SA 라이선스 서버 변경                           ==========")
        Console.WriteLine("========== 20180702 - CATIA GLOBAL 라이선스 서버 추가 (172.100.100.65) =========")
        Console.WriteLine("========== 20180702 - CATIA 한국 샤시 서버 추가 (유럽사무소 이관)     ==========")
        Console.WriteLine("========== 20180717 - CATIA GLOBAL 라이선스 서버 추가 (210.105.188.203) ========")
        Console.WriteLine("========== 20190131 - 3DCS(SHB) 라이선스 서버 추가 (172.100.100.64)     ========")
        Console.WriteLine("========== 20190710 - 3DCS(중국상해) 라이선스 서버 추가 (172.100.100.66)========")
        Console.WriteLine("========== 20200306 - 해석 Fluent 서버 수정(203.250.10.249)             ========")
        Console.WriteLine("========== 20200306 - 전자 ANSYS 서버 추가 (172.20.100.114)             ========")
        Console.WriteLine("========== 20200309 - 해석 DAFUL 서버 추가 (203.250.10.249)             ========")
        Console.WriteLine("========== 20200309 - 해석 STARCCM 서버 추가 (172.100.100.65)           ========")
        Console.WriteLine("========== 20200309 - 해석 OASYS 서버 추가 (203.250.10.249)             ========")
        Console.WriteLine("========== 20200310 - 해석 SHERLOCK 서버 추가 (203.250.10.249)          ========")
        Console.WriteLine("========== 20200311 - 해석 CRADLE 서버 추가 (203.250.10.59)             ========")
        Console.WriteLine("========== 20200323 - 금형 MOLDFLOW 서버 추가 (203.250.10.25)           ========")
        Console.WriteLine("========== 20200525 - 전자 CST 서버 추가 (172.20.100.114)               ========")
        Console.WriteLine("========== 20200617 - 금형 MOLDFLOW 서버 삭제 (203.250.10.25):해석 서버로 통합==")
        Console.WriteLine("========== 20200525 - 해석 CST 서버 추가 (203.250.10.50)                ========")

        Dim strCategory As String = "Restart"
        oSysOp.SendMail(strCategory, "License Monitoring System Restart")

        'D========================================================================
        'Dim MSSqlConn As New SqlConnection
        'MSSqlConn = _DB._CAD_SQLConnect()

        'Dim SQLComm As New SqlCommand
        'Try
        '    SQLComm.Connection = MSSqlConn
        'Catch ex As Exception
        '    Console.WriteLine("     >>SQLCommand ERROR!!!")
        '    Exit Sub
        'End Try
        'Dim tempDate As New Date(2018, 8, 23, 12, 0, 0)
        ''_DB.InsertData(tempDate)
        'Dim TEST_DS As New DataSet


        ''_DB.UserInformation_Sync(TEST_DS, MSSqlConn, SQLComm)
        '_DB.Interface_UserInfo(MSSqlConn, SQLComm, tempDate)
        ''============================================================================

        Dim CheckCount As Integer = 0

        SearchTime = New Thread(AddressOf GetTime)
        SearchTime.Start()

        '========== 라이선스 정보 테이블 생성 ==========
        Dim DS As New DataSet
        Dim Lic_SrvDT As DataTable
        Lic_SrvDT = DS.Tables.Add("LicSrvDT")
        Lic_SrvDT = _DB.GetLic_SRV(DS, Lic_SrvDT)

        Dim Lic_InfoDT As DataTable
        Lic_InfoDT = DS.Tables.Add("Lic_InfoDT")
        '==============================================

        Do
            Dim CheckTime As Date = My.Computer.Clock.LocalTime
            Console.WriteLine("[IntervalTime : 10초.....]")

            Thread.Sleep(IntervalTime)

            Console.WriteLine(">>>>> 서버에서 라이선스 정보를 확인 합니다.(" & CheckCount & " : " & My.Computer.Clock.LocalTime & ") <<<<<")
            Lic_InfoDT = _DB.GetLic_Info(DS, Lic_InfoDT)

            If Lic_InfoDT IsNot Nothing Or Lic_InfoDT.Rows.Count <> 0 Then
                Console.WriteLine(">>>>> DSLS 라이선스를 체크합니다(" & CheckCount & " : " & My.Computer.Clock.LocalTime & ") <<<<<")

                Dim arrDT1 As New ArrayList
                arrDT1 = _CATIA.CheckStart(Lic_InfoDT, Lic_SrvDT)

                Console.WriteLine(">>>>> LMTOOLS 라이선스를 체크합니다(" & CheckCount & " : " & My.Computer.Clock.LocalTime & ") <<<<<")
                Dim arrDT2 As New ArrayList
                arrDT2 = _NX.CheckStart(arrDT1(0), arrDT1(1))

                Console.WriteLine(">>>>> 라이선스 정보를 업데이트 합니다.(" & CheckCount & " : " & My.Computer.Clock.LocalTime & ") <<<<<")

                ''=====================로컬 테스트=======================
                'arrDT2.Add(Lic_InfoDT)
                'arrDT2.Add(Lic_SrvDT)
                ''=======================================================

                _DB.DBUpdate(DS, arrDT2(0), arrDT2(1))  '정상 체크 시 사용
                '_DB.DBUpdate(DS, arrDT1(0), arrDT1(1)) 'CATIA 체크 없이 테스트 시 사용

                Lic_SrvDT.Clear()
                Lic_InfoDT.Rows.Clear()
                Lic_InfoDT.Columns.Clear()
            End If

            If InsertTime = True Then
                Console.WriteLine("======= 라이선스 모니터링 DB 업데이트 시작 : " & InsertDate & " =======")
                _DB.InsertData(InsertDate)
                Console.WriteLine("======= 라이선스 모니터링 업데이트 완료 : " & InsertDate & " =======")
                InsertTime = False
            End If
            CheckCount = CheckCount + 1

        Loop
    End Sub

    Private Sub GetTime()

        On Error Resume Next

        Do
            Thread.Sleep(60000)
            If InsertTime = False Then
                If My.Computer.Clock.LocalTime.Minute = 0 Then
                    Console.WriteLine(" >>>>> 라이선스 정보 DB 업데이트 확인 : " & My.Computer.Clock.LocalTime & " <<<<<")
                    Console.WriteLine(" >>>>> 라이선스 정보 DB 입력 시간 확인 : " & My.Computer.Clock.LocalTime.Minute & " <<<<<")
                    InsertDate = My.Computer.Clock.LocalTime
                    InsertTime = True
                Else
                    InsertTime = False
                End If
            End If
        Loop

    End Sub



End Module
