
Imports System.IO
Imports System.Data.SqlClient


'=======================테이블 정보(중요,StartMonitoring 에서 테이블 변경시 반드시 변경해줄것)======================
'CADLicDT.Columns.Add("CAD_TYPE")        '0
'CADLicDT.Columns.Add("LIC_TYPE")        '1
'CADLicDT.Columns.Add("LIC_VER")         '2
'CADLicDT.Columns.Add("LIC_SRV")         '3
'CADLicDT.Columns.Add("USER_ID")         '4
'CADLicDT.Columns.Add("NAME")            '5
'CADLicDT.Columns.Add("DEPT_ID")         '6
'CADLicDT.Columns.Add("DEPT_NAME")       '7
'CADLicDT.Columns.Add("IP_INFO")         '8
'CADLicDT.Columns.Add("BIZ_CD")          '9
'CADLicDT.Columns.Add("M_ORG_NM")        '10
'CADLicDT.Columns.Add("C_ORG_CD")        '11
'CADLicDT.Columns.Add("C_DRG_NM")        '12
'CADLicDT.Columns.Add("LIC_NATION")      '13
'====================================================================================================================

Public Class CATIALicSrvCheck

    Dim oSysOp As New SysOperation

    Public Function CheckStart(ByVal Lic_InfoDT As DataTable, ByVal Lic_SrvDT As DataTable) As ArrayList

        Dim arrLicInfo As New ArrayList

        Console.WriteLine("1. CATIA 라이선스 확인 중 : " & My.Computer.Clock.LocalTime)

        '연구소라이선스 확인
        '172.100.100.61 LICSRV-KOR-RND
        Dim SrvInfo As String = String.Empty

        SrvInfo = "172.100.100.61(LICSRV - 연구소1)"
        Console.WriteLine("   1-1) 172.100.100.61 - 연구소1 : " & My.Computer.Clock.LocalTime)
        arrLicInfo = oSysOp.DSLSLicInfo_NEW(arrLicInfo, "172.100.100.61")
        Lic_SrvDT = GetDSLSLicenseServer_Information(Lic_InfoDT, Lic_SrvDT, arrLicInfo, "172.100.100.61")

        '172.100.100.62 LICSRV-KOR-RND
        arrLicInfo.Clear()
        SrvInfo = "172.100.100.62(LICSRV - 연구소2)"
        Console.WriteLine("   1-2) 172.100.100.62 - 연구소2 : " & My.Computer.Clock.LocalTime)
        arrLicInfo = oSysOp.DSLSLicInfo_NEW(arrLicInfo, "172.100.100.62")
        Lic_SrvDT = GetDSLSLicenseServer_Information(Lic_InfoDT, Lic_SrvDT, arrLicInfo, "172.100.100.62")

        '172.100.100.63 LICSRV-KOR-RND
        arrLicInfo.Clear()
        SrvInfo = "172.100.100.63(LICSRV - 연구소3)"
        Console.WriteLine("   1-3) 172.100.100.63 - 연구소3 : " & My.Computer.Clock.LocalTime)
        arrLicInfo = oSysOp.DSLSLicInfo_NEW(arrLicInfo, "172.100.100.63")
        Lic_SrvDT = GetDSLSLicenseServer_Information(Lic_InfoDT, Lic_SrvDT, arrLicInfo, "172.100.100.63")

        '172.100.100.64 LICSRV-KOR-RND
        arrLicInfo.Clear()
        SrvInfo = "172.100.100.64(LICSRV - 연구소4)"
        Console.WriteLine("   1-4) 172.100.100.64 - 연구소4 : " & My.Computer.Clock.LocalTime)
        arrLicInfo = oSysOp.DSLSLicInfo_NEW(arrLicInfo, "172.100.100.64")
        Lic_SrvDT = GetDSLSLicenseServer_Information(Lic_InfoDT, Lic_SrvDT, arrLicInfo, "172.100.100.64")

        '172.100.100.66 LICSRV-CHN-ENG
        arrLicInfo.Clear()
        SrvInfo = "172.100.100.66(LICSRV - 중국(ENG) CATIA)"
        Console.WriteLine("   1-5) 172.100.100.66 - 중국 ENG : " & My.Computer.Clock.LocalTime)
        arrLicInfo = oSysOp.DSLSLicInfo_NEW(arrLicInfo, "172.100.100.66")
        Lic_SrvDT = GetDSLSLicenseServer_Information(Lic_InfoDT, Lic_SrvDT, arrLicInfo, "172.100.100.66")

        '172.100.100.67	LICSRV-PDM
        arrLicInfo.Clear()
        SrvInfo = "172.100.100.67(LICSRV - PDM: CT5, DCI)"
        Console.WriteLine("   1-6) 172.100.100.67 - PDM : " & My.Computer.Clock.LocalTime)
        arrLicInfo = oSysOp.DSLSLicInfo_NEW(arrLicInfo, "172.100.100.67")
        Lic_SrvDT = GetDSLSLicenseServer_Information(Lic_InfoDT, Lic_SrvDT, arrLicInfo, "172.100.100.67")

        ''172.20.100.67 LICSRV-KOR-SUP
        'arrLicInfo.Clear()
        'SrvInfo = "172.20.100.67(LICSRV - 협력사 APR)"
        'Console.WriteLine("   1-7) 172.20.100.67 - 협력사 APR : " & My.Computer.Clock.LocalTime)
        'arrLicInfo = oSysOp.DSLSLicInfo_NEW(arrLicInfo, "172.20.100.67")
        'Lic_SrvDT = GetDSLSLicenseServer_Information(Lic_InfoDT, Lic_SrvDT, arrLicInfo, "172.20.100.67")

        '172.20.100.70	LICSRV-KOR-ENG
        arrLicInfo.Clear()
        SrvInfo = "172.20.100.70(LICSRV - 국내 APR)"
        Console.WriteLine("   1-7) 172.20.100.70 - 국내 APR : " & My.Computer.Clock.LocalTime)
        arrLicInfo = oSysOp.DSLSLicInfo_NEW(arrLicInfo, "172.20.100.70")
        Lic_SrvDT = GetDSLSLicenseServer_Information(Lic_InfoDT, Lic_SrvDT, arrLicInfo, "172.20.100.70")

        '172.20.100.71	LICSRV-CHN-BK
        arrLicInfo.Clear()
        SrvInfo = "172.20.100.71(LICSRV - 중국(BK) CATIA, APR)"
        Console.WriteLine("   1-8) 172.20.100.71 - 중국(BK) : " & My.Computer.Clock.LocalTime)
        arrLicInfo = oSysOp.DSLSLicInfo_NEW(arrLicInfo, "172.20.100.71")
        Lic_SrvDT = GetDSLSLicenseServer_Information(Lic_InfoDT, Lic_SrvDT, arrLicInfo, "172.20.100.71")

        '172.20.100.72	LICSRV-CHN-DP
        arrLicInfo.Clear()
        SrvInfo = "172.20.100.72(LICSRV - 중국(DP) CATIA, APR)"
        Console.WriteLine("   1-9) 172.20.100.72 - 중국(DP) : " & My.Computer.Clock.LocalTime)
        arrLicInfo = oSysOp.DSLSLicInfo_NEW(arrLicInfo, "172.20.100.72")
        Lic_SrvDT = GetDSLSLicenseServer_Information(Lic_InfoDT, Lic_SrvDT, arrLicInfo, "172.20.100.72")

        '172.20.100.73	LICSRV-CHN-SS
        arrLicInfo.Clear()
        SrvInfo = "172.20.100.73(LICSRV - 중국(SS) CATIA, APR)"
        Console.WriteLine("   1-10) 172.20.100.73 - 중국(SS) : " & My.Computer.Clock.LocalTime)
        arrLicInfo = oSysOp.DSLSLicInfo_NEW(arrLicInfo, "172.20.100.73")
        Lic_SrvDT = GetDSLSLicenseServer_Information(Lic_InfoDT, Lic_SrvDT, arrLicInfo, "172.20.100.73")

        '172.20.100.74	LICSRV-CHN-YD
        arrLicInfo.Clear()
        SrvInfo = "172.20.100.74(LICSRV - 중국(YD) CATIA, APR)"
        Console.WriteLine("   1-11) 172.20.100.74 - 중국(YD) : " & My.Computer.Clock.LocalTime)
        arrLicInfo = oSysOp.DSLSLicInfo_NEW(arrLicInfo, "172.20.100.74")
        Lic_SrvDT = GetDSLSLicenseServer_Information(Lic_InfoDT, Lic_SrvDT, arrLicInfo, "172.20.100.74")

        '172.20.100.87	LICSRV-CHN-CQ
        arrLicInfo.Clear()
        SrvInfo = "172.20.100.87(LICSRV - 중국(CQ) CATIA)"
        Console.WriteLine("   1-12) 172.20.100.87 - 중국(CQ) : " & My.Computer.Clock.LocalTime)
        arrLicInfo = oSysOp.DSLSLicInfo_NEW(arrLicInfo, "172.20.100.87")
        Lic_SrvDT = GetDSLSLicenseServer_Information(Lic_InfoDT, Lic_SrvDT, arrLicInfo, "172.20.100.87")

        '172.20.100.75	LICSRV-IND
        arrLicInfo.Clear()
        SrvInfo = "172.20.100.75(LICSRV - 인도 CATIA, APR)"
        Console.WriteLine("   1-13) 172.100.100.221 - 인도 : " & My.Computer.Clock.LocalTime)
        arrLicInfo = oSysOp.DSLSLicInfo_NEW(arrLicInfo, "172.100.100.221")
        Lic_SrvDT = GetDSLSLicenseServer_Information(Lic_InfoDT, Lic_SrvDT, arrLicInfo, "172.100.100.221")

        '172.20.100.76	LICSRV-USA-TN_L
        arrLicInfo.Clear()
        SrvInfo = "172.20.100.76(LICSRV - 미국-테네시램프 APR)"
        Console.WriteLine("   1-14) 172.20.100.76 - 미국 테네시 램프 : " & My.Computer.Clock.LocalTime)
        arrLicInfo = oSysOp.DSLSLicInfo_NEW(arrLicInfo, "172.20.100.76")
        Lic_SrvDT = GetDSLSLicenseServer_Information(Lic_InfoDT, Lic_SrvDT, arrLicInfo, "172.20.100.76")

        '172.20.100.77	LICSRV-USA-TN_C
        arrLicInfo.Clear()
        SrvInfo = "172.20.100.77(LICSRV - 미국-테네시샤시 APR)"
        Console.WriteLine("   1-15) 172.20.100.77 - 미국 테네시 샤시 : " & My.Computer.Clock.LocalTime)
        arrLicInfo = oSysOp.DSLSLicInfo_NEW(arrLicInfo, "172.20.100.77")
        Lic_SrvDT = GetDSLSLicenseServer_Information(Lic_InfoDT, Lic_SrvDT, arrLicInfo, "172.20.100.77")

        '172.20.100.78	LICSRV-USA-AL
        arrLicInfo.Clear()
        SrvInfo = "172.20.100.78(LICSRV - 미국-알라바마 APR)"
        Console.WriteLine("   1-16) 172.20.100.78 - 미국 알라바마 : " & My.Computer.Clock.LocalTime)
        arrLicInfo = oSysOp.DSLSLicInfo_NEW(arrLicInfo, "172.20.100.78")
        Lic_SrvDT = GetDSLSLicenseServer_Information(Lic_InfoDT, Lic_SrvDT, arrLicInfo, "172.20.100.78")

        '172.20.100.80	LICSRV-POL
        arrLicInfo.Clear()
        SrvInfo = "172.20.100.80(LICSRV - 폴란드 APR)"
        Console.WriteLine("   1-17) 172.20.100.80 - 폴란드 : " & My.Computer.Clock.LocalTime)
        arrLicInfo = oSysOp.DSLSLicInfo_NEW(arrLicInfo, "172.20.100.80")
        Lic_SrvDT = GetDSLSLicenseServer_Information(Lic_InfoDT, Lic_SrvDT, arrLicInfo, "172.20.100.80")

        ''203.250.10.160	LICSRV-KOR-MOLD
        'arrLicInfo.Clear()
        'SrvInfo = "203.250.13.160(LICSRV - 금형 CATIA)"
        'Console.WriteLine("   1-18) 203.250.13.160 - 금형 : " & My.Computer.Clock.LocalTime)
        'arrLicInfo = oSysOp.DSLSLicInfo_NEW(arrLicInfo, "203.250.13.160")
        'Lic_SrvDT = GetDSLSLicenseServer_Information(Lic_InfoDT, Lic_SrvDT, arrLicInfo, "203.250.13.160")

        '203.250.10.25	LICSRV-KOR-MOLD
        arrLicInfo.Clear()
        SrvInfo = "203.250.10.22(LICSRV - 금형 CATIA)"
        Console.WriteLine("   1-18) 203.250.10.22 - 금형 : " & My.Computer.Clock.LocalTime)
        arrLicInfo = oSysOp.DSLSLicInfo_NEW(arrLicInfo, "203.250.10.22")
        Lic_SrvDT = GetDSLSLicenseServer_Information(Lic_InfoDT, Lic_SrvDT, arrLicInfo, "203.250.10.22")

        '203.250.92.10	    LICSRV-USA-ENG
        arrLicInfo.Clear()
        SrvInfo = "203.250.92.10(LICSRV - 미국-ENG CATIA) 확인"
        Console.WriteLine("   1-19) 203.250.92.10 - 미국 ENG : " & My.Computer.Clock.LocalTime)
        arrLicInfo = oSysOp.DSLSLicInfo_NEW(arrLicInfo, "203.250.92.10")
        Lic_SrvDT = GetDSLSLicenseServer_Information(Lic_InfoDT, Lic_SrvDT, arrLicInfo, "203.250.92.10")

        '172.20.100.115	LICSRV-KOR-INJ
        arrLicInfo.Clear()
        SrvInfo = "172.20.100.115(LICSRV - 사출표면 CATIA)"
        Console.WriteLine("   1-20) 210.105.188.201 - 사출표면 : " & My.Computer.Clock.LocalTime)
        arrLicInfo = oSysOp.DSLSLicInfo_NEW(arrLicInfo, "172.20.100.115")
        Lic_SrvDT = GetDSLSLicenseServer_Information(Lic_InfoDT, Lic_SrvDT, arrLicInfo, "172.20.100.115")

        '172.20.100.116	LICSRV-KOR-SA
        arrLicInfo.Clear()
        SrvInfo = "172.20.100.116(LICSRV - 시스템자동화 CATIA)"
        Console.WriteLine("   1-21) 172.20.100.116 - 시스템자동화 : " & My.Computer.Clock.LocalTime)
        arrLicInfo = oSysOp.DSLSLicInfo_NEW(arrLicInfo, "172.20.100.116")
        Lic_SrvDT = GetDSLSLicenseServer_Information(Lic_InfoDT, Lic_SrvDT, arrLicInfo, "172.20.100.116")

        '203.250.10.25    LICSRV-KOR-MOLD V6
        arrLicInfo.Clear()
        SrvInfo = "203.250.10.25(LICSRV - 금형V6)"
        Console.WriteLine("   1-22) 203.250.10.25 - 금형V6 : " & My.Computer.Clock.LocalTime)
        arrLicInfo = oSysOp.DSLSLicInfo_NEW(arrLicInfo, "203.250.10.25")
        Lic_SrvDT = GetDSLSLicenseServer_Information(Lic_InfoDT, Lic_SrvDT, arrLicInfo, "203.250.10.25")

        '172.100.100.65    LICSRV-GLOBAL(IND)
        arrLicInfo.Clear()
        SrvInfo = "172.100.100.65(LICSRV - GLOBAL)"
        Console.WriteLine("   1-23) 172.100.100.65 - GLOBAL : " & My.Computer.Clock.LocalTime)
        arrLicInfo = oSysOp.DSLSLicInfo_NEW(arrLicInfo, "172.100.100.65")
        Lic_SrvDT = GetDSLSLicenseServer_Information(Lic_InfoDT, Lic_SrvDT, arrLicInfo, "172.100.100.65")


        '210.105.188.203    LICSRV-GLOBAL(KOR)
        arrLicInfo.Clear()
        SrvInfo = "210.105.188.203(LICSRV - GLOBAL)"
        Console.WriteLine("   1-24) 210.105.188.203 - GLOBAL : " & My.Computer.Clock.LocalTime)
        arrLicInfo = oSysOp.DSLSLicInfo_NEW(arrLicInfo, "210.105.188.203")
        Lic_SrvDT = GetDSLSLicenseServer_Information(Lic_InfoDT, Lic_SrvDT, arrLicInfo, "210.105.188.203")

        '172.20.100.118    LICSRV-KOR(CHASSIS)
        arrLicInfo.Clear()
        SrvInfo = "172.20.100.118(LICSRV - KOR-샤시)"
        Console.WriteLine("   1-25) 172.20.100.118 - KOR-샤시 : " & My.Computer.Clock.LocalTime)
        arrLicInfo = oSysOp.DSLSLicInfo_NEW(arrLicInfo, "172.20.100.118")
        Lic_SrvDT = GetDSLSLicenseServer_Information(Lic_InfoDT, Lic_SrvDT, arrLicInfo, "172.20.100.118")

        Dim arrDT As New ArrayList
        arrDT.Add(Lic_InfoDT)
        arrDT.Add(Lic_SrvDT)

        Return arrDT

    End Function

    Private Function GetCATIALicenseServer_Information(ByVal Lic_InfoDT As DataTable, ByVal Lic_SrvDT As DataTable, ByVal arrlicInfo As ArrayList, ByVal strIP As String) As DataTable

        For index As Integer = 8 To arrlicInfo.Count - 1
            Debug.Print(arrlicInfo.Item(index))
            If arrlicInfo.Item(index).ToString.Contains("Product Name:") Then

                Dim arrValue As New ArrayList
                Dim CadType As String = Nothing
                Dim Lictype As String = Nothing
                Dim LicVer As String = Nothing
                Dim LicNum As Integer = 0
                Dim LicOnline As Integer = 0
                Dim LicOffline As Integer = 0
                Dim UserID As String = Nothing
                Dim UserIP As String = Nothing
                Dim LicSrv As String = strIP
                Dim LicNa As String = Nothing
                Dim LicDate As String = oSysOp.SortDateTime(My.Computer.Clock.LocalTime)
                Dim LIcPro As String = Nothing

                arrValue = oSysOp.GetValue(arrlicInfo.Item(index).ToString.Split(" "))
                Lictype = arrValue(arrValue.Count - 1)

                If Lictype <> "Product" Then
                    arrValue = oSysOp.GetValue(arrlicInfo.Item(index + 1).ToString.Split(" "))
                    LicVer = arrValue(arrValue.Count - 1)
                    Dim GetLicRow As DataRow()
                    GetLicRow = Lic_InfoDT.Select("LIC_VER='" & LicVer & "' AND LIC_TYPE='" & Lictype & "' AND LIC_SRV='" & LicSrv & "'")

                    If GetLicRow.Count <> 0 Then
                        CadType = GetLicRow(0).Item("CAD_TYPE")
                        arrValue = oSysOp.GetValue(arrlicInfo.Item(index + 10).ToString.Split(" "))
                        LicNum = arrValue(0)
                        LicOnline = arrValue(1)
                        LicOffline = arrValue(2)
                        LicSrv = GetLicRow(0).Item("LIC_SRV")
                        LicNa = GetLicRow(0).Item("LIC_NATION")
                        LIcPro = GetLicRow(0).Item("LIC_PROPERTY")

                        GetLicRow(0).Item("LIC_NUM") = LicNum
                        GetLicRow(0).Item("LIC_ONLINE") = LicOnline
                        GetLicRow(0).Item("LIC_OFFLINE") = LicOffline
                    End If


                    Try
                        Dim iUser As Integer = index + 14
                        For index1 As Integer = 0 To LicOnline - 1
                            arrValue = oSysOp.GetValue(arrlicInfo.Item(iUser).ToString.Split(" "))
                            UserID = arrValue(1)
                            arrValue = oSysOp.GetValue(arrlicInfo.Item(iUser + 3).ToString.Split(" "))
                            UserIP = ConvertIP(arrValue(1))
                            iUser = iUser + 5

                            Dim NewRow As DataRow
                            NewRow = Lic_SrvDT.Rows.Add()
                            NewRow.Item("CAD_TYPE") = CadType
                            NewRow.Item("LIC_TYPE") = Lictype
                            NewRow.Item("LIC_VER") = LicVer
                            NewRow.Item("LIC_SRV") = LicSrv
                            NewRow.Item("USER_ID") = UserID
                            NewRow.Item("IP_INFO") = UserIP
                            NewRow.Item("LIC_NATION") = LicNa
                            NewRow.Item("LIC_DATE") = LicDate
                            NewRow.Item("LIC_PROPERTY") = LIcPro
                        Next

                        For index1 As Integer = 0 To LicOffline - 1
                            Dim NewRow As DataRow
                            NewRow = Lic_SrvDT.Rows.Add()
                            NewRow.Item("CAD_TYPE") = CadType
                            NewRow.Item("LIC_TYPE") = Lictype
                            NewRow.Item("LIC_VER") = LicVer
                            NewRow.Item("LIC_SRV") = LicSrv
                            NewRow.Item("USER_ID") = "OFFLINE_USER"
                            NewRow.Item("IP_INFO") = "-"
                            NewRow.Item("LIC_NATION") = LicNa
                            NewRow.Item("LIC_DATE") = LicDate
                            NewRow.Item("LIC_PROPERTY") = LIcPro
                        Next

                    Catch ex As Exception
                        Console.WriteLine(Lictype & "/" & LicVer & "/" & LicSrv & " : Error!!!")
                        Console.WriteLine("ERROR :" & ex.Message)
                    End Try
                End If
            End If
        Next

        Return Lic_SrvDT

    End Function


    Private Function ConvertIP(ByVal strIP As String) As String

        Dim GetIP() As String
        Dim hexIP As String

        Dim IP1 As String
        Dim IP2 As String
        Dim IP3 As String
        Dim IP4 As String
        Dim IP As String

        Try
            GetIP = strIP.Split(".")
            hexIP = GetIP(0)
            IP1 = Val("&H0" & (hexIP.ElementAt(0) & hexIP.ElementAt(1)))
            IP2 = Val("&H0" & (hexIP.ElementAt(2) & hexIP.ElementAt(3)))
            IP3 = Val("&H0" & (hexIP.ElementAt(4) & hexIP.ElementAt(5)))
            IP4 = Val("&H0" & (hexIP.ElementAt(6) & hexIP.ElementAt(7)))
            IP = IP1 & "." & IP2 & "." & IP3 & "." & IP4
        Catch ex As Exception
            IP = "000.000.000.000"
        End Try


        Return IP

    End Function

    Private Function GetDSLSLicenseServer_Information_Before(ByVal Lic_InfoDT As DataTable, ByVal Lic_SrvDT As DataTable, ByVal arrlicInfo As ArrayList, ByVal strIP As String) As DataTable
        Dim arrValue As New ArrayList
        Dim LicType As String
        Dim arrLicType As New ArrayList

        Dim LicNum As Integer
        Dim arrLicNum As New ArrayList

        Dim LicOnline As Integer
        Dim arrLicOnline As New ArrayList

        Dim LicOffline As Integer = 0
        Dim arrLicOffline As New ArrayList

        Dim UserID As String = Nothing
        Dim UserIP As String = Nothing
        Dim LicSrv As String = strIP
        Dim strUserIP() As String = Nothing
        Dim LicVer As String = "DSLS"

        Dim CADType As String = Nothing
        Dim LicNa As String = Nothing
        Dim LicPro As String = Nothing
        Dim LicDate As String = oSysOp.SortDateTime(My.Computer.Clock.LocalTime)
        Dim intIndexLictype As Integer
        Dim NewRow As DataRow




        For index1 As Integer = 7 To arrlicInfo.Count - 1   'TEXT 전체 검색

            If arrlicInfo.Item(index1).ToString.Contains("maxReleaseNumber") Then
                Dim GetLicRow As DataRow()
                arrValue = oSysOp.GetValue(arrlicInfo.Item(index1).ToString.Split(" "))
                LicType = arrValue(0).ToString.Remove(0, 2)
                GetLicRow = Lic_InfoDT.Select("LIC_TYPE='" & LicType & "' AND LIC_SRV='" & LicSrv & "'")
                If GetLicRow.Count <> 0 Then
                    CADType = GetLicRow(0).Item("CAD_TYPE")
                    LicSrv = GetLicRow(0).Item("LIC_SRV")
                    LicNa = GetLicRow(0).Item("LIC_NATION")
                    LicPro = GetLicRow(0).Item("LIC_PROPERTY")
                End If
                Try
                    If arrLicType.Contains(LicType) = False Then    '동일한 라이선스 타입이 포함되어 있지 않은 경우
                        arrLicType.Add(LicType)
                        LicNum = CType(arrValue(18).ToString, Integer)
                        arrLicNum.Add(LicNum)
                        LicOnline = CType(arrValue(20).ToString, Integer)
                        arrLicOnline.Add(LicOnline)

                        If LicOnline <> 0 Then  '사용자 정보 추출
                            For index2 As Integer = index1 + 1 To index1 + (LicOnline * 2)  'For index2 As Integer = index1 + 1 To index1 + (LicOnline * 2)
                                Try
                                    If arrlicInfo.Item(index2).ToString.Contains("internal Id") Then
                                        arrValue = oSysOp.GetValue(arrlicInfo.Item(index2).ToString.Split(" "))
                                        UserID = arrValue.Item(20).ToString
                                        strUserIP = arrValue.Item(arrValue.Count - 1).ToString.Split("/")
                                        If arrlicInfo.Item(index2 + 1).ToString.Contains("offline licenseId") Then
                                            UserIP = "OFFLINE_USER"
                                            LicOffline = LicOffline + 1
                                        Else
                                            UserIP = strUserIP(1).ToString
                                        End If
                                        NewRow = Lic_SrvDT.Rows.Add()
                                        NewRow.Item("CAD_TYPE") = CADType
                                        NewRow.Item("LIC_TYPE") = LicType
                                        NewRow.Item("LIC_VER") = LicVer
                                        NewRow.Item("LIC_SRV") = LicSrv
                                        NewRow.Item("USER_ID") = UserID
                                        NewRow.Item("IP_INFO") = UserIP
                                        NewRow.Item("LIC_NATION") = LicNa
                                        NewRow.Item("LIC_DATE") = LicDate
                                        NewRow.Item("LIC_PROPERTY") = LicPro
                                    End If
                                Catch ex As Exception
                                    Debug.Print(ex.ToString)
                                End Try
                            Next
                            Debug.Print("TEST")
                            Try
                                arrLicOffline.Add(LicOffline)
                            Catch ex As Exception
                                ex.ToString()
                            End Try
                        Else
                            arrLicOffline.Add(0)
                        End If

                    Else                                            '동일한 라이선스 타입이 포함되어 있는 경우
                        LicNum = CType(arrValue(18).ToString, Integer)
                        LicOnline = CType(arrValue(20).ToString, Integer)
                        intIndexLictype = arrLicType.IndexOf(LicType)
                        arrLicNum(intIndexLictype) = arrLicNum(intIndexLictype) + LicNum
                        arrLicOnline(intIndexLictype) = arrLicOnline(intIndexLictype) + LicOnline
                        If LicOnline <> 0 Then  '사용자 정보 추출
                            For index2 As Integer = index1 + 1 To index1 + (LicOnline * 2)  'For index2 As Integer = index1 + 1 To index1 + (LicOnline * 2)
                                If arrlicInfo.Item(index2).ToString.Contains("internal Id") Then
                                    arrValue = oSysOp.GetValue(arrlicInfo.Item(index2).ToString.Split(" "))
                                    UserID = arrValue.Item(20).ToString
                                    strUserIP = arrValue.Item(arrValue.Count - 1).ToString.Split("/")
                                    If arrlicInfo.Item(index2 + 1).ToString.Contains("offline licenseId") Then
                                        UserIP = "OFFLINE_USER"
                                        arrLicOffline(intIndexLictype) = arrLicOffline(intIndexLictype) + 1
                                    Else
                                        UserIP = strUserIP(1).ToString
                                    End If
                                    NewRow = Lic_SrvDT.Rows.Add()
                                    NewRow.Item("CAD_TYPE") = CADType
                                    NewRow.Item("LIC_TYPE") = LicType
                                    NewRow.Item("LIC_VER") = LicVer
                                    NewRow.Item("LIC_SRV") = LicSrv
                                    NewRow.Item("USER_ID") = UserID
                                    NewRow.Item("IP_INFO") = UserIP
                                    NewRow.Item("LIC_NATION") = LicNa
                                    NewRow.Item("LIC_DATE") = LicDate
                                    NewRow.Item("LIC_PROPERTY") = LicPro
                                End If
                            Next
                        End If
                    End If
                Catch ex As Exception

                End Try

            End If
        Next

        For icount As Integer = 0 To arrLicType.Count - 1
            Dim GetLicRow As DataRow()
            GetLicRow = Lic_InfoDT.Select("LIC_TYPE='" & arrLicType(icount).ToString & "' AND LIC_SRV='" & LicSrv & "'")
            If GetLicRow.Count <> 0 Then
                GetLicRow(0).Item("LIC_NUM") = arrLicNum(icount)
                GetLicRow(0).Item("LIC_ONLINE") = arrLicOnline(icount)
                GetLicRow(0).Item("LIC_OFFLINE") = arrLicOffline(icount)
            Else
                Console.WriteLine("LIC 정보와 미일치:" & LicSrv & "/" & arrLicType(icount).ToString & "/" & My.Computer.Clock.LocalTime)
            End If
        Next

        Debug.Print("Complete")
        Return Lic_SrvDT
    End Function

    Private Function GetDSLSLicenseServer_Information(ByVal Lic_InfoDT As DataTable, ByVal Lic_SrvDT As DataTable, ByVal arrlicInfo As ArrayList, ByVal strIP As String) As DataTable
        Dim arrValue As New ArrayList
        Dim LicType As String
        Dim arrLicType As New ArrayList

        Dim LicNum As Integer
        Dim arrLicNum As New ArrayList

        Dim LicOnline As Integer
        Dim arrLicOnline As New ArrayList


        Dim arrLicOffline As New ArrayList

        Dim UserID As String = Nothing
        Dim UserIP As String = Nothing
        Dim LicSrv As String = strIP
        Dim strUserIP() As String = Nothing
        Dim LicVer As String = "DSLS"

        Dim CADType As String = Nothing
        Dim LicNa As String = Nothing
        Dim LicPro As String = Nothing
        Dim LicDate As String = oSysOp.SortDateTime(My.Computer.Clock.LocalTime)
        Dim intIndexLictype As Integer
        Dim NewRow As DataRow

        For index1 As Integer = 7 To arrlicInfo.Count - 1   'TEXT 전체 검색

            If arrlicInfo.Item(index1).ToString.Contains("maxReleaseNumber") Then
                Dim GetLicRow As DataRow()
                arrValue = oSysOp.GetValue(arrlicInfo.Item(index1).ToString.Split(" "))
                LicType = arrValue(0).ToString.Remove(0, 2)
                GetLicRow = Lic_InfoDT.Select("LIC_TYPE='" & LicType & "' AND LIC_SRV='" & LicSrv & "'")
                If GetLicRow.Count <> 0 Then
                    CADType = GetLicRow(0).Item("CAD_TYPE")
                    LicSrv = GetLicRow(0).Item("LIC_SRV")
                    LicNa = GetLicRow(0).Item("LIC_NATION")
                    LicPro = GetLicRow(0).Item("LIC_PROPERTY")
                End If

                If arrLicType.Contains(LicType) = False Then    '동일한 라이선스 타입이 포함되어 있지 않은 경우
                    arrLicType.Add(LicType)
                    LicNum = CType(arrValue(18).ToString, Integer)
                    arrLicNum.Add(LicNum)
                    LicOnline = CType(arrValue(20).ToString, Integer)
                    arrLicOnline.Add(LicOnline)

                    If LicOnline <> 0 Then  '사용자 정보 추출
                        Dim indexforUser As Integer = index1
                        Dim bolRowCount1 As Boolean = False
                        Dim bolRowCount2 As Boolean = False

                        Dim LicOffline As Integer = 0
                        Do
                            If arrlicInfo.Item(indexforUser).ToString.Contains("internal Id") Then
                                arrValue = oSysOp.GetValue(arrlicInfo.Item(indexforUser).ToString.Split(" "))
                                Try
                                    UserID = arrValue.Item(20).ToString
                                Catch ex As Exception
                                    UserID = "-"
                                End Try


                                Try
                                    'strUserIP = arrValue.Item(arrValue.Count - 1).ToString.Split("/")
                                    strUserIP = arrValue.Item(24).ToString.Split("/")
                                    If arrlicInfo.Item(indexforUser + 1).ToString.Contains("offline licenseId") Then
                                        UserIP = "OFFLINE_USER"
                                        LicOffline = LicOffline + 1
                                    Else
                                        UserIP = strUserIP(1).ToString
                                    End If
                                Catch ex As Exception
                                    UserIP = "-"
                                End Try
                                Try
                                    NewRow = Lic_SrvDT.Rows.Add()
                                    NewRow.Item("CAD_TYPE") = CADType
                                    NewRow.Item("LIC_TYPE") = LicType
                                    NewRow.Item("LIC_VER") = LicVer
                                    NewRow.Item("LIC_SRV") = LicSrv
                                    NewRow.Item("USER_ID") = UserID
                                    NewRow.Item("IP_INFO") = UserIP
                                    NewRow.Item("LIC_NATION") = LicNa
                                    NewRow.Item("LIC_DATE") = LicDate
                                    NewRow.Item("LIC_PROPERTY") = LicPro
                                Catch ex As Exception
                                    Debug.Print(ex.ToString)
                                End Try
                            End If
                            indexforUser = indexforUser + 1
                            bolRowCount1 = arrlicInfo.Item(indexforUser).ToString.Contains("maxReleaseNumber")
                            If indexforUser = arrlicInfo.Count - 1 Then
                                bolRowCount2 = True
                            End If
                        Loop Until bolRowCount1 = True Or bolRowCount2 = True
                        arrLicOffline.Add(LicOffline)
                    Else
                        arrLicOffline.Add(0)
                    End If
                Else                                            '동일한 라이선스 타입이 포함되어 있는 경우
                    LicNum = CType(arrValue(18).ToString, Integer)
                    LicOnline = CType(arrValue(20).ToString, Integer)
                    intIndexLictype = arrLicType.IndexOf(LicType)
                    arrLicNum(intIndexLictype) = arrLicNum(intIndexLictype) + LicNum
                    arrLicOnline(intIndexLictype) = arrLicOnline(intIndexLictype) + LicOnline
                    If LicOnline <> 0 Then  '사용자 정보 추출
                        Dim indexforUser As Integer = index1
                        Dim bolRowCount3 As Boolean = False
                        Dim bolRowCount4 As Boolean = False
                        Do
                            If arrlicInfo.Item(indexforUser).ToString.Contains("internal Id") Then
                                arrValue = oSysOp.GetValue(arrlicInfo.Item(indexforUser).ToString.Split(" "))
                                Try
                                    UserID = arrValue.Item(20).ToString
                                Catch ex As Exception
                                    UserID = "-"
                                End Try


                                Try
                                    'strUserIP = arrValue.Item(arrValue.Count - 1).ToString.Split("/")
                                    strUserIP = arrValue.Item(24).ToString.Split("/")
                                    If arrlicInfo.Item(indexforUser + 1).ToString.Contains("offline licenseId") Then
                                        UserIP = "OFFLINE_USER"
                                        'LicOffline = LicOffline + 1
                                        arrLicOffline(intIndexLictype) = arrLicOffline(intIndexLictype) + 1
                                    Else
                                        UserIP = strUserIP(1).ToString
                                    End If
                                Catch ex As Exception
                                    UserIP = "-"
                                End Try
                                Try
                                    NewRow = Lic_SrvDT.Rows.Add()
                                    NewRow.Item("CAD_TYPE") = CADType
                                    NewRow.Item("LIC_TYPE") = LicType
                                    NewRow.Item("LIC_VER") = LicVer
                                    NewRow.Item("LIC_SRV") = LicSrv
                                    NewRow.Item("USER_ID") = UserID
                                    NewRow.Item("IP_INFO") = UserIP
                                    NewRow.Item("LIC_NATION") = LicNa
                                    NewRow.Item("LIC_DATE") = LicDate
                                    NewRow.Item("LIC_PROPERTY") = LicPro
                                Catch ex As Exception
                                    Debug.Print(ex.ToString)
                                End Try

                            End If
                            indexforUser = indexforUser + 1
                            bolRowCount3 = arrlicInfo.Item(indexforUser).ToString.Contains("maxReleaseNumber")
                            If indexforUser = arrlicInfo.Count - 1 Then
                                bolRowCount4 = True
                            End If

                        Loop Until bolRowCount3 = True Or bolRowCount4 = True
                        'arrlicInfo.Item(indexforUser).ToString.Contains("maxReleaseNumber") Or arrlicInfo.Count - 1
                    End If
                End If
            End If


        Next

        For icount As Integer = 0 To arrLicType.Count - 1
            Dim GetLicRow As DataRow()
            GetLicRow = Lic_InfoDT.Select("LIC_TYPE='" & arrLicType(icount).ToString & "' AND LIC_SRV='" & LicSrv & "'")
            If GetLicRow.Count <> 0 Then
                GetLicRow(0).Item("LIC_NUM") = arrLicNum(icount)
                GetLicRow(0).Item("LIC_ONLINE") = arrLicOnline(icount) - arrLicOffline(icount)
                GetLicRow(0).Item("LIC_OFFLINE") = arrLicOffline(icount)
            Else
                Console.WriteLine("LIC 정보와 미일치:" & LicSrv & "/" & arrLicType(icount).ToString & "/" & My.Computer.Clock.LocalTime)
            End If
        Next

        Debug.Print("Complete")
        Return Lic_SrvDT
    End Function
End Class
