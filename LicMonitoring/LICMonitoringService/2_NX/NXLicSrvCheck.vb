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
'====================================================================================================================

Public Class NXLicSrvCheck

    Dim SysOp As New SysOperation
    Dim DBConn As New DBConnect
    Dim DS As New DataSet
    Dim lmutillPath As String = "C:\Program Files\Siemens\PLMLicenseServer\lmutil.exe" 'C:\DCS
    'Dim lmutillPath As String = "C:\eng_apps\PLMLicenseServer\lmutil.exe" 'C:\DCS
    Dim altairutilPath As String = "C:\Program Files\Altair\licensing12.2\bin\lmxendutil.exe"
    Dim flexnetPath As String = "C:\Program Files\CD-adapco\FLEXlm\11_14_0_2\bin\lmutil.exe"

    Public Function CheckStart(ByVal Lic_InfoDT As DataTable, ByVal Lic_SrvDT As DataTable) As ArrayList
        Dim arrNXLicInfo As New ArrayList
        Dim SrvInfo As String = String.Empty
        Console.WriteLine("1. NX 라이선스 확인 중 : " & My.Computer.Clock.LocalTime)

        '연구소 NX 라이선스 확인
        SrvInfo = "NX 연구소 라이선스(172.100.100.65)"
        Dim NXArg1 As String = "lmstat -c 28000@172.100.100.65 -a"
        arrNXLicInfo = SysOp.NXLicInfo_NEW(arrNXLicInfo, lmutillPath, NXArg1, SrvInfo)
        Lic_SrvDT = GetNXLicenseServer_Information(Lic_InfoDT, Lic_SrvDT, arrNXLicInfo, "172.100.100.65", "COMMON")
        Console.WriteLine("   1-1) NX 연구소 라이선스 확인 완료(172.100.100.65) : " & My.Computer.Clock.LocalTime)

        '글로벌 라이선스 확인 (한국/유럽)
        arrNXLicInfo.Clear()
        SrvInfo = "NX 글로벌 라이선스 (172.100.100.66)"
        Dim NXArg2 As String = "lmstat -c 28000@172.100.100.66 -a"
        arrNXLicInfo = SysOp.NXLicInfo_NEW(arrNXLicInfo, lmutillPath, NXArg2, SrvInfo)
        Lic_SrvDT = GetNXLicenseServer_Information(Lic_InfoDT, Lic_SrvDT, arrNXLicInfo, "172.100.100.66", "COMMON")
        Console.WriteLine("   1-2) NX 글로벌 라이선스(한국/유럽) 확인 완료(172.100.100.66) : " & My.Computer.Clock.LocalTime)

        '글로벌 라이선스 확인 (중국)
        arrNXLicInfo.Clear()
        SrvInfo = "NX 글로벌 라이선스 (172.20.100.67)"
        Dim NXArg3 As String = "lmstat -c 28000@172.20.100.67 -a"
        arrNXLicInfo = SysOp.NXLicInfo_NEW(arrNXLicInfo, lmutillPath, NXArg3, SrvInfo)
        Lic_SrvDT = GetNXLicenseServer_Information(Lic_InfoDT, Lic_SrvDT, arrNXLicInfo, "172.20.100.67", "COMMON")
        Console.WriteLine("   1-3) NX 글로벌 라이선스(중국) 확인 완료(172.20.100.67) : " & My.Computer.Clock.LocalTime)

        '중국 현지 라이선스 확인
        arrNXLicInfo.Clear()
        SrvInfo = "NX 중국 로컬 라이선스 (192.168.53.221)"
        Dim NXArg4 As String = "lmstat -c 28000@192.168.53.221 -a"
        arrNXLicInfo = SysOp.NXLicInfo_NEW(arrNXLicInfo, lmutillPath, NXArg4, SrvInfo)
        Lic_SrvDT = GetNXLicenseServer_Information(Lic_InfoDT, Lic_SrvDT, arrNXLicInfo, "192.168.53.221", "COMMON")
        Console.WriteLine("   1-4) NX 중국 라이선스 확인 완료(192.168.53.221) : " & My.Computer.Clock.LocalTime)

        '미국 현지 라이선스 확인
        arrNXLicInfo.Clear()
        SrvInfo = "NX 미국 라이선스 (203.250.92.10)"
        Dim NXArg5 As String = "lmstat -c 28000@203.250.92.10 -a"
        arrNXLicInfo = SysOp.NXLicInfo_NEW(arrNXLicInfo, lmutillPath, NXArg5, SrvInfo)
        Lic_SrvDT = GetNXLicenseServer_Information(Lic_InfoDT, Lic_SrvDT, arrNXLicInfo, "203.250.92.10", "COMMON")
        Console.WriteLine("   1-5) NX 미국 라이선스 확인 완료(203.250.92.10) : " & My.Computer.Clock.LocalTime)

        Console.WriteLine("2. 3DCS 라이선스 확인 중 : " & My.Computer.Clock.LocalTime)

        arrNXLicInfo.Clear()
        SrvInfo = "3DCS 라이선스 (172.100.100.62)"
        Dim DCSArg1 As String = "lmstat -c 27000@172.100.100.62 -a"
        arrNXLicInfo = SysOp.NXLicInfo_NEW(arrNXLicInfo, lmutillPath, DCSArg1, SrvInfo)
        Lic_SrvDT = GetNXLicenseServer_Information(Lic_InfoDT, Lic_SrvDT, arrNXLicInfo, "172.100.100.62", "ADD")
        Console.WriteLine("   2-1) 3DCS 라이선스 확인 완료(172.100.100.62) : " & My.Computer.Clock.LocalTime)

        arrNXLicInfo.Clear()
        SrvInfo = "3DCS 라이선스 (172.100.100.63)"
        Dim DCSArg2 As String = "lmstat -c 27000@172.100.100.63 -a"
        arrNXLicInfo = SysOp.NXLicInfo_NEW(arrNXLicInfo, lmutillPath, DCSArg2, SrvInfo)
        Lic_SrvDT = GetNXLicenseServer_Information(Lic_InfoDT, Lic_SrvDT, arrNXLicInfo, "172.100.100.63", "ADD")
        Console.WriteLine("   2-2) 3DCS 라이선스 확인 완료(172.100.100.63) : " & My.Computer.Clock.LocalTime)

        arrNXLicInfo.Clear()
        SrvInfo = "3DCS 라이선스 (172.100.100.64)"
        Dim DCSArg3 As String = "lmstat -c 27000@172.100.100.64 -a"
        arrNXLicInfo = SysOp.NXLicInfo_NEW(arrNXLicInfo, lmutillPath, DCSArg3, SrvInfo)
        Lic_SrvDT = GetNXLicenseServer_Information(Lic_InfoDT, Lic_SrvDT, arrNXLicInfo, "172.100.100.64", "ADD")
        Console.WriteLine("   2-3) 3DCS 라이선스 확인 완료(172.100.100.64) : " & My.Computer.Clock.LocalTime)

        arrNXLicInfo.Clear()
        SrvInfo = "3DCS 라이선스 (172.100.100.66)"
        Dim DCSArg4 As String = "lmstat -c 27000@172.100.100.66 -a"
        arrNXLicInfo = SysOp.NXLicInfo_NEW(arrNXLicInfo, lmutillPath, DCSArg4, SrvInfo)
        Lic_SrvDT = GetNXLicenseServer_Information(Lic_InfoDT, Lic_SrvDT, arrNXLicInfo, "172.100.100.66", "ADD")
        Console.WriteLine("   2-4) 3DCS 라이선스 확인 완료(172.100.100.66) : " & My.Computer.Clock.LocalTime)

        Console.WriteLine("3. Alias 라이선스 확인 중 : " & My.Computer.Clock.LocalTime)

        arrNXLicInfo.Clear()
        SrvInfo = "Alias 라이선스 (172.100.100.67)"
        Dim AliasArg1 As String = "lmstat -c 27000@172.100.100.67 -a"
        arrNXLicInfo = SysOp.NXLicInfo_NEW(arrNXLicInfo, lmutillPath, AliasArg1, SrvInfo)
        Lic_SrvDT = GetNXLicenseServer_Information(Lic_InfoDT, Lic_SrvDT, arrNXLicInfo, "172.100.100.67", "ADD")
        Console.WriteLine("   3-1) Alias 라이선스 확인 완료 : " & My.Computer.Clock.LocalTime)

        Console.WriteLine("4. ALM 라이선스 확인 중 : " & My.Computer.Clock.LocalTime)

        arrNXLicInfo.Clear()
        SrvInfo = "Integrity 라이선스 (172.100.100.125)"
        Dim ALMArg1 As String = "lmstat -c 27000@172.100.100.125 -a"
        arrNXLicInfo = SysOp.NXLicInfo_NEW(arrNXLicInfo, lmutillPath, ALMArg1, SrvInfo)
        Lic_SrvDT = GetNXLicenseServer_Information(Lic_InfoDT, Lic_SrvDT, arrNXLicInfo, "172.100.100.125", "ADD")
        Console.WriteLine("   4-1) Integrity 라이선스 확인 완료 : " & My.Computer.Clock.LocalTime)

        arrNXLicInfo.Clear()
        SrvInfo = "IBM TOOL 라이선스 (192.168.10.12)"
        Dim ALMArg2 As String = "lmstat -c 27000@192.168.10.12 -a"
        arrNXLicInfo = SysOp.NXLicInfo_NEW(arrNXLicInfo, lmutillPath, ALMArg2, SrvInfo)
        Lic_SrvDT = GetNXLicenseServer_Information(Lic_InfoDT, Lic_SrvDT, arrNXLicInfo, "192.168.10.12", "COMMON")
        Console.WriteLine("   4-2) IBM TOOL 라이선스 확인 완료 : " & My.Computer.Clock.LocalTime)

        Console.WriteLine("5. ALLEGRO 라이선스 확인 중 : " & My.Computer.Clock.LocalTime)

        arrNXLicInfo.Clear()
        SrvInfo = "ALLEGRO 라이선스 (192.168.10.9)"
        Dim ALLEGROArg1 As String = "lmstat -c 5280@192.168.10.9 -a"
        arrNXLicInfo = SysOp.NXLicInfo_NEW(arrNXLicInfo, lmutillPath, ALLEGROArg1, SrvInfo)
        Lic_SrvDT = GetNXLicenseServer_Information(Lic_InfoDT, Lic_SrvDT, arrNXLicInfo, "192.168.10.9", "ADD")
        Console.WriteLine("   5-1) ALLEGRO 라이선스 확인 완료 : " & My.Computer.Clock.LocalTime)

        Console.WriteLine("6. MATLAB 라이선스 확인 중 : " & My.Computer.Clock.LocalTime)

        arrNXLicInfo.Clear()
        SrvInfo = "MATRLAB 라이선스 (172.20.32.50)"
        Dim MATLABArg1 As String = "lmstat -c 27001@172.20.32.50 -a"
        arrNXLicInfo = SysOp.NXLicInfo_NEW(arrNXLicInfo, lmutillPath, MATLABArg1, SrvInfo)
        Lic_SrvDT = GetNXLicenseServer_Information(Lic_InfoDT, Lic_SrvDT, arrNXLicInfo, "172.20.32.50", "ADD")
        Console.WriteLine("   6) MATLAB 라이선스 확인 완료 : " & My.Computer.Clock.LocalTime)

        Console.WriteLine("7. MOLDFLOW 라이선스 확인 중 : " & My.Computer.Clock.LocalTime)

        arrNXLicInfo.Clear()
        SrvInfo = "MOLDFLOW 라이선스 (203.250.10.249)"
        Dim MOLDFLOWArg1 As String = "lmstat -c 27003@203.250.10.249 -a"
        arrNXLicInfo = SysOp.NXLicInfo_NEW(arrNXLicInfo, lmutillPath, MOLDFLOWArg1, SrvInfo)
        'Lic_SrvDT = GetNXLicenseServer_Information(Lic_InfoDT, Lic_SrvDT, arrNXLicInfo, "203.250.10.249", "ADD")
        Lic_SrvDT = GetMOLDFLOW_Information(Lic_InfoDT, Lic_SrvDT, arrNXLicInfo, "203.250.10.249", "ADD")
        Console.WriteLine("   7-1) MOLDFLOW 한국 라이선스 확인 완료 : " & My.Computer.Clock.LocalTime)

        arrNXLicInfo.Clear()
        SrvInfo = "MOLDFLOW 라이선스 (192.168.9.140)"
        Dim MOLDFLOWArg2 As String = "lmstat -c 27000@192.168.9.140 -a"
        arrNXLicInfo = SysOp.NXLicInfo_NEW(arrNXLicInfo, lmutillPath, MOLDFLOWArg2, SrvInfo)
        'Lic_SrvDT = GetNXLicenseServer_Information(Lic_InfoDT, Lic_SrvDT, arrNXLicInfo, "203.250.10.249", "ADD")
        Lic_SrvDT = GetMOLDFLOW_Information(Lic_InfoDT, Lic_SrvDT, arrNXLicInfo, "192.168.9.140", "ADD")
        Console.WriteLine("   7-2) MOLDFLOW 인도 라이선스 확인 완료 : " & My.Computer.Clock.LocalTime)

        'arrNXLicInfo.Clear()
        'SrvInfo = "MOLDFLOW 라이선스 (203.250.10.25)"
        'Dim MOLDFLOWArg3 As String = "lmstat -c 27000@203.250.10.25 -a"
        'arrNXLicInfo = SysOp.NXLicInfo_NEW(arrNXLicInfo, lmutillPath, MOLDFLOWArg3, SrvInfo)
        ''Lic_SrvDT = GetNXLicenseServer_Information(Lic_InfoDT, Lic_SrvDT, arrNXLicInfo, "203.250.10.249", "ADD")
        'Lic_SrvDT = GetMOLDFLOW_Information(Lic_InfoDT, Lic_SrvDT, arrNXLicInfo, "203.250.10.25", "ADD")
        'Console.WriteLine("   7-3) MOLDFLOW 금형 라이선스 확인 완료 : " & My.Computer.Clock.LocalTime)

        Console.WriteLine("8. CODESCROLL 라이선스 확인 중 : " & My.Computer.Clock.LocalTime)

        arrNXLicInfo.Clear()
        SrvInfo = "CODESCROLL 라이선스 (172.20.33.105)"
        Dim CODESCROLLArg1 As String = "lmstat -c 27010@172.20.33.105 -a"
        arrNXLicInfo = SysOp.NXLicInfo_NEW(arrNXLicInfo, lmutillPath, CODESCROLLArg1, SrvInfo)
        Lic_SrvDT = GetNXLicenseServer_Information(Lic_InfoDT, Lic_SrvDT, arrNXLicInfo, "172.20.33.105", "COMMON")
        Console.WriteLine("   8) CODESCROLL 라이선스 확인 완료 : " & My.Computer.Clock.LocalTime)

        Console.WriteLine("9. FLUENT 라이선스 확인 중 : " & My.Computer.Clock.LocalTime)

        arrNXLicInfo.Clear()
        SrvInfo = "FLUENT 라이선스(203.250.10.48)"
        Dim FLUENTArg As String = "lmstat -c 1055@203.250.10.48 -a"
        arrNXLicInfo = SysOp.NXLicInfo_NEW(arrNXLicInfo, lmutillPath, FLUENTArg, SrvInfo)
        Lic_SrvDT = GetFlexLm_DoLoop_LicenseServer_Information(Lic_InfoDT, Lic_SrvDT, arrNXLicInfo, "203.250.10.48", "FLUENT")
        Console.WriteLine("   9) FLUENT 라이선스 확인 완료(203.250.10.48) : " & My.Computer.Clock.LocalTime)

        Console.WriteLine("10. HYPERWORKS 라이선스 확인 중 : " & My.Computer.Clock.LocalTime)

        arrNXLicInfo.Clear()
        SrvInfo = "HYPERWORKS 한국 라이선스(203.250.10.249)"
        Dim HWArg As String = "-licstat -host 203.250.10.249 -port 6200"
        arrNXLicInfo = SysOp.NXLicInfo_NEW(arrNXLicInfo, altairutilPath, HWArg, SrvInfo)
        Lic_SrvDT = GetHyperWorksLicenseServer_Information(Lic_InfoDT, Lic_SrvDT, arrNXLicInfo, "203.250.10.249", "KOR")
        Console.WriteLine("   10-1) HYPERWORKS 한국 라이선스 확인 완료(203.250.10.249) : " & My.Computer.Clock.LocalTime)

        arrNXLicInfo.Clear()
        SrvInfo = "HYPERWORKS 인도 라이선스(192.168.8.17)"
        Dim HW_IND_Arg As String = "-licstat -host 192.168.8.17 -port 6200"
        arrNXLicInfo = SysOp.NXLicInfo_NEW(arrNXLicInfo, altairutilPath, HW_IND_Arg, SrvInfo)
        Lic_SrvDT = GetHyperWorksLicenseServer_Information(Lic_InfoDT, Lic_SrvDT, arrNXLicInfo, "192.168.8.17", "IND")
        Console.WriteLine("   10-2) HYPERWORKS 인도 라이선스 확인 완료(192.168.8.17) : " & My.Computer.Clock.LocalTime)

        Console.WriteLine("11. ABAQUS 라이선스 확인 중 : " & My.Computer.Clock.LocalTime)

        arrNXLicInfo.Clear()
        SrvInfo = "ABAQUS 라이선스(203.250.10.249)"
        Dim ABAQUSArg As String = "lmstat -c 27000@203.250.10.249 -a"
        arrNXLicInfo = SysOp.NXLicInfo_NEW(arrNXLicInfo, lmutillPath, ABAQUSArg, SrvInfo)
        Lic_SrvDT = GetABAQUSLicenseServer_Information(Lic_InfoDT, Lic_SrvDT, arrNXLicInfo, "203.250.10.249")
        Console.WriteLine("   11) ABAQUS 라이선스 확인 완료(203.250.10.249) : " & My.Computer.Clock.LocalTime)

        Console.WriteLine("12. ANSYS 라이선스 확인 중 : " & My.Computer.Clock.LocalTime)

        arrNXLicInfo.Clear()
        SrvInfo = "ANSYS 라이선스(172.20.100.114)"
        Dim ANSYSArg As String = "lmstat -c 1055@172.20.100.114 -a"
        arrNXLicInfo = SysOp.NXLicInfo_NEW(arrNXLicInfo, flexnetPath, ANSYSArg, SrvInfo)
        Lic_SrvDT = GetFlexLm_DoLoop_LicenseServer_Information(Lic_InfoDT, Lic_SrvDT, arrNXLicInfo, "172.20.100.114", "ANSYS")
        Console.WriteLine("   12) ANSYS 라이선스 확인 완료(172.20.100.114) : " & My.Computer.Clock.LocalTime)

        Console.WriteLine("13. DAFUL 라이선스 확인 중 : " & My.Computer.Clock.LocalTime)

        arrNXLicInfo.Clear()
        SrvInfo = "DAFUL 라이선스(203.250.10.249)"
        Dim DAFULArg As String = "lmstat -c 27001@203.250.10.249 -a"
        arrNXLicInfo = SysOp.NXLicInfo_NEW(arrNXLicInfo, lmutillPath, DAFULArg, SrvInfo)
        Lic_SrvDT = GetFlexLm_NoVersion_LicenseServer_Information(Lic_InfoDT, Lic_SrvDT, arrNXLicInfo, "203.250.10.249", "COMMON")
        Console.WriteLine("   13) DAFUL 라이선스 확인 완료(203.250.10.249) : " & My.Computer.Clock.LocalTime)

        Console.WriteLine("14. STARCCM 라이선스 확인 중 : " & My.Computer.Clock.LocalTime)

        arrNXLicInfo.Clear()
        SrvInfo = "STARCCM 라이선스(172.100.100.65)"
        Dim STARCCMArg As String = "lmstat -c 1999@172.100.100.65 -a"
        arrNXLicInfo = SysOp.NXLicInfo_NEW(arrNXLicInfo, lmutillPath, STARCCMArg, SrvInfo)
        Lic_SrvDT = GetFlexLm_NoVersion_LicenseServer_Information(Lic_InfoDT, Lic_SrvDT, arrNXLicInfo, "172.100.100.65", "COMMON")
        Console.WriteLine("   14) STARCCM 라이선스 확인 완료(172.100.100.65) : " & My.Computer.Clock.LocalTime)

        Console.WriteLine("15. OASYS 라이선스 확인 중 : " & My.Computer.Clock.LocalTime)

        arrNXLicInfo.Clear()
        SrvInfo = "OASYS 라이선스(203.250.10.249)"
        Dim OASYSArg As String = "lmstat -c 28001@203.250.10.249 -a"
        arrNXLicInfo = SysOp.NXLicInfo_NEW(arrNXLicInfo, lmutillPath, OASYSArg, SrvInfo)
        Lic_SrvDT = GetFlexLm_NoVersion_LicenseServer_Information(Lic_InfoDT, Lic_SrvDT, arrNXLicInfo, "203.250.10.249", "ADD")
        Console.WriteLine("   15) OASYS 라이선스 확인 완료(203.250.10.249) : " & My.Computer.Clock.LocalTime)

        Console.WriteLine("16. SHERLOCK 라이선스 확인 중 : " & My.Computer.Clock.LocalTime)

        arrNXLicInfo.Clear()
        SrvInfo = "SHERLOCK 라이선스(203.250.10.249)"
        Dim SHERLOCKArg As String = "lmstat -c 27002@203.250.10.249 -a"
        arrNXLicInfo = SysOp.NXLicInfo_NEW(arrNXLicInfo, flexnetPath, SHERLOCKArg, SrvInfo)
        Lic_SrvDT = GetFlexLm_NoVersion_LicenseServer_Information(Lic_InfoDT, Lic_SrvDT, arrNXLicInfo, "203.250.10.249", "ADD")
        Console.WriteLine("   16) SHERLOCK 라이선스 확인 완료(203.250.10.249) : " & My.Computer.Clock.LocalTime)

        Console.WriteLine("17. CRADLE 라이선스 확인 중 : " & My.Computer.Clock.LocalTime)

        arrNXLicInfo.Clear()
        SrvInfo = "CRADLE 라이선스(203.250.10.59)"
        Dim CREADLEArg As String = "lmstat -c 32646@203.250.10.59 -a"
        arrNXLicInfo = SysOp.NXLicInfo_NEW(arrNXLicInfo, lmutillPath, CREADLEArg, SrvInfo)
        Lic_SrvDT = GetFlexLm_DoLoop_LicenseServer_Information(Lic_InfoDT, Lic_SrvDT, arrNXLicInfo, "203.250.10.59", "CRADLE")
        Console.WriteLine("   17) CRADLE 라이선스 확인 완료(203.250.10.59) : " & My.Computer.Clock.LocalTime)

        Console.WriteLine("18. CST 라이선스 확인 중 : " & My.Computer.Clock.LocalTime)

        arrNXLicInfo.Clear()
        SrvInfo = "CST 라이선스(172.20.100.114)"
        Dim CSTArg1 As String = "lmstat -c 27000@172.20.100.114 -a"
        arrNXLicInfo = SysOp.NXLicInfo_NEW(arrNXLicInfo, lmutillPath, CSTArg1, SrvInfo)
        Lic_SrvDT = GetFlexLm_NoVersion_LicenseServer_Information(Lic_InfoDT, Lic_SrvDT, arrNXLicInfo, "172.20.100.114", "ADD")
        Console.WriteLine("   18-1) CST(전자) 라이선스 확인 완료(172.20.100.114) : " & My.Computer.Clock.LocalTime)

        arrNXLicInfo.Clear()
        SrvInfo = "CST 라이선스(203.250.10.50)"
        Dim CSTArg2 As String = "lmstat -c 27100@203.250.10.50 -a"
        arrNXLicInfo = SysOp.NXLicInfo_NEW(arrNXLicInfo, lmutillPath, CSTArg2, SrvInfo)
        Lic_SrvDT = GetFlexLm_NoVersion_LicenseServer_Information(Lic_InfoDT, Lic_SrvDT, arrNXLicInfo, "203.250.10.50", "ADD")
        Console.WriteLine("   18-2) CST(해석) 라이선스 확인 완료(203.250.10.50) : " & My.Computer.Clock.LocalTime)

        Dim arrDT As New ArrayList

        arrDT.Add(Lic_InfoDT)
        arrDT.Add(Lic_SrvDT)

        Return arrDT

    End Function

    Private Function GetNXLicenseServer_Information(ByVal Lic_InfoDT As DataTable, ByVal Lic_SrvDT As DataTable, ByVal arrNXLicInfo As ArrayList, ByVal strIP As String, ByVal strCAD As String) As DataTable

        Dim iNX13100N As Integer = 0
        Dim arrLicSrvInfo As New ArrayList
        Dim FilterStr As String = "Users of "
        Dim strFilterforOffline As String = "(linger:"

        For index As Integer = 11 To arrNXLicInfo.Count - 1
            Debug.Print(arrNXLicInfo.Item(index))
            If arrNXLicInfo.Item(index).ToString.Contains(FilterStr) Then
                Dim arrValue As New ArrayList
                Dim CadType As String
                Dim Lictype As String
                Dim LicVer As String = "-"
                Dim LicNum As Integer
                Dim LicOnline As Integer
                Dim UserID As String
                Dim UserIP As String = "-"
                Dim LicSrv As String = strIP
                Dim LicNa As String
                Dim LicDate As String = SysOp.SortDateTime(My.Computer.Clock.LocalTime)
                Dim LicPro As String
                Dim LicOffline As Integer = 0

                arrValue = SysOp.GetValue(arrNXLicInfo.Item(index).ToString.Split(" "))
                Lictype = arrValue(2).ToString.Replace(":", "")

                Try
                    LicNum = arrValue(5)
                Catch ex As Exception
                    LicNum = LicNum
                End Try

                Try
                    LicOnline = arrValue(10)
                Catch ex As Exception
                    LicOnline = LicOnline
                End Try

                'If LicOnline <> 0 Then
                '    arrValue = SysOp.GetValue(arrNXLicInfo.Item(index + 2).ToString.Split(" "))
                '    LicVer = arrValue(1).ToString.Replace(",", "")
                'Else
                '    'LMTOOLS 는 사용자가 없을 경우 버전 및 전체 사용자가 나오지 않아서 사용자가 없는 경우를 셋팅해줘야함
                '    Select Case Lictype
                '        Case "GMS4050"
                '            If LicSrv = "172.100.100.65" Then
                '                LicVer = "v29.0"
                '            ElseIf LicSrv = "192.168.53.221" Then
                '                LicVer = "v29.0"
                '            Else
                '                LicVer = "v30.0"
                '            End If

                '        Case "MFG"
                '            LicVer = "v35.0"
                '        Case "NX13100N"
                '            If LicSrv = "172.100.100.65" Then
                '                LicVer = "v35.0"
                '            ElseIf LicSrv = "172.100.100.66" Then
                '                LicVer = "v35.0"
                '            ElseIf LicSrv = "203.250.92.10" Then
                '                LicVer = "v35.0"
                '            ElseIf LicSrv = "192.168.53.221" Then
                '                LicVer = "v29.0"
                '            ElseIf LicSrv = "172.20.100.67" Then
                '                LicVer = "v32.0"
                '            Else
                '                LicVer = "v30.0"
                '            End If
                '        Case "catv5_nx_sca"
                '            LicVer = "v35.0"
                '            'If LicSrv = "172.100.100.66" Then
                '            '    LicVer = "v35.0"
                '            'Else
                '            '    LicVer = "v34.0"
                '            'End If
                '        Case "DCS_3DCS_CAT"
                '            LicVer = "v7.6"
                '            'ALIAS
                '        Case "85300SURFST_F"
                '            LicVer = "v1.000"
                '        Case "85300SURFST_T_F"
                '            LicVer = "v1.000"
                '            'VRED
                '        Case "86307VRDDES_T_F"
                '            LicVer = "v1.000"
                '        Case "DCS_3DCS_ANALYST_SA_S"
                '            LicVer = "v7.6"
                '        Case "MKSIntegrityManager-Float"
                '            LicVer = "v5.0"
                '        Case "MKSSourceIntegrity-Float"
                '            LicVer = "v5.0"
                '        Case "DOORS"
                '            LicVer = "v2014.0630"
                '        Case "ClearQuest"
                '            LicVer = "v1.10000"
                '        Case "ClearCase"
                '            LicVer = "v1.00000"
                '            'Case "AllegroSigrity_PI_Base"
                '            '    LicVer = "v2016.0"
                '            'Case "AllegroSigrity_PI_Signoff_Opt"
                '            '    LicVer = "v2016.0"
                '            'Case "AllegroSigrity_SI_Base"
                '            '    LicVer = "v2016.0"
                '        Case "Allegro_Auth_HighSpeed_Option"
                '            LicVer = "v17.4"
                '        Case "Allegro_Design_Publisher"
                '            LicVer = "v17.4"
                '            'Case "Allegro_Viewer_Plus"  '제외
                '            '    LicVer = "v17.2"
                '            'Case "Allegro_performance"  '제외
                '            '    LicVer = "v17.2"
                '            'Case "Allegro_studio"   '제외
                '            '    LicVer = "v17.2"
                '            'Case "Base_Verilog_Lib" '제외
                '            '    LicVer = "v17.2"
                '        Case "Concept_HDL_studio"
                '            LicVer = "v17.4"
                '            'Case "Extended_Verilog_Lib" '제외
                '            '    LicVer = "v17.2"
                '        Case "PCB_design_studio"
                '            LicVer = "v17.4"
                '        Case "PCB_librarian_expert"
                '            LicVer = "v17.4"
                '        Case "PSpiceStudio"
                '            LicVer = "v17.4"
                '            'Case "SPECCTRA_HP"  '제외
                '            '    LicVer = "v17.2"
                '            'Case "SPECCTRA_PCB" '제외
                '            '    LicVer = "v17.2"
                '            'Case "expgen"   '제외
                '            '    LicVer = "v17.2"
                '            'Case "pcomp"    '제외
                '            '    LicVer = "v17.2"
                '            'Case "plotVersa"    '제외
                '            '    LicVer = "v17.2"
                '        Case "MATLAB"
                '            LicVer = "v42"
                '        Case "SIMULINK"
                '            LicVer = "v42"
                '        Case "Video_and_Image_Blockset"
                '            LicVer = "v42"
                '        Case "Control_Toolbox"
                '            LicVer = "v42"
                '        Case "Signal_Blocks"
                '            LicVer = "v42"
                '        Case "RTW_Embedded_Coder"
                '            LicVer = "v42"
                '        Case "Fixed_Point_Toolbox"
                '            LicVer = "v42"
                '        Case "Image_Acquisition_Toolbox"
                '            LicVer = "v42"
                '        Case "Image_Toolbox"
                '            LicVer = "v42"
                '        Case "MATLAB_Coder"
                '            LicVer = "v42"
                '        Case "MATLAB_Report_Gen"
                '            LicVer = "v42"
                '        Case "Signal_Toolbox"
                '            LicVer = "v42"
                '        Case "Real-Time_Workshop"
                '            LicVer = "v42"
                '        Case "Simulink_Design_Verifier"
                '            LicVer = "v42"
                '        Case "SIMULINK_Report_Gen"
                '            LicVer = "v42"
                '        Case "Simulink_Test"
                '            LicVer = "v42"
                '        Case "SL_Verification_Validation"
                '            LicVer = "v42"
                '        Case "Stateflow"
                '            LicVer = "v42"
                '        Case "Statistics_Toolbox"
                '            LicVer = "v42"
                '        Case "SystemTest"
                '            LicVer = "v42"
                '        Case "Fixed-Point_Blocks"
                '            LicVer = "v42"
                '        Case "Filter_Design_Toolbox"
                '            LicVer = "v42"
                '        Case "Stateflow_Coder"
                '            LicVer = "v42"

                '            'MoldFlow
                '        Case "77800MFS_T_F"
                '            LicVer = "v1.000"
                '        Case "86802PLC0000023_T_F"
                '            LicVer = "v1.000"
                '        Case "86387MFIP_T_F"
                '            LicVer = "v1.000"
                '            'Case "86586MFIB_F"
                '            '    LicVer = "v1.000"
                '        Case "85816MFAA_2012_0F"
                '            LicVer = "v1.000"
                '        Case "85819MFAM_2012_0F"
                '            LicVer = "v1.000"
                '        Case "86751MFIB_2017_0F"
                '            LicVer = "v1.000"
                '            'Case "86556MFIB_2016_0F"
                '            '    LicVer = "v1.000"
                '            'Case "86324MFIB_2015_0F"
                '            '    LicVer = "v1.000"
                '            'Case "86157MFIB_2014_0F"
                '            '    LicVer = "v1.000"
                '        Case "86752MFIP_2017_0F"
                '            LicVer = "v1.000"
                '        Case "86555MFIP_2016_0F"
                '            LicVer = "v1.000"
                '        Case "86323MFIP_2015_0F"
                '            LicVer = "v1.000"
                '        Case "86155MFIP_2014_0F"
                '            LicVer = "v1.000"
                '        Case "86802PLC0000023_2017_0F"
                '            LicVer = "v1.000"
                '        Case "86754MFS_2017_0F"
                '            LicVer = "v1.000"
                '        Case "86558MFS_2016_0F"
                '            LicVer = "v1.000"
                '        Case "86326MFS_2015_0F"
                '            LicVer = "v1.000"
                '        Case "86158MFS_2014_0F"
                '            LicVer = "v1.000"
                '        Case "87198MFIA_2019_0F"
                '            LicVer = "v1.000"
                '        Case "87016MFIA_2018_0F"
                '            LicVer = "v1.000"
                '        Case "86753MFIA_2017_0F"
                '            LicVer = "v1.000"
                '        Case "86557MFIA_2016_0F"
                '            LicVer = "v1.000"
                '        Case "86325MFIA_2015_0F"
                '            LicVer = "v1.000"
                '        Case "86156MFIA_2014_0F"
                '            LicVer = "v1.000"
                '        Case "87196MFIP_2019_0F"
                '            LicVer = "v1.000"
                '        Case "87014MFIP_2018_0F"
                '            LicVer = "v1.000"
                '        Case "87032PLC0000023_2018_0F"
                '            LicVer = "v1.000"
                '        Case "87194MFS_2019_0F"
                '            LicVer = "v1.000"
                '        Case "87017MFS_2018_0F"
                '            LicVer = "v1.000"

                '        Case "csct"
                '            LicVer = "v2.1"
                '        Case "full"
                '            LicVer = "v2.1"
                '        Case "FTC_csct"
                '            LicVer = "v2.1"
                '        Case "FTCP_csct"
                '            LicVer = "v2.1"
                '    End Select
                'End If

                '1열 NX, 2열 3DCS, 3열 Alias/VRED, 4열 ALM, 5~7열 ALLEGRO, 8~10열 MATLAB, 11~14열 MOLDFLOW, 15열 CODESCROLL
                If Lictype = "GMS4050" Or Lictype = "MFG" Or Lictype = "NX13100N" Or _
                Lictype = "catv5_nx_sca" Or Lictype = "DCS_3DCS_CAT" Or Lictype = "DCS_3DCS_ANALYST_SA_S" Or _
                Lictype = "85300SURFST_F" Or Lictype = "85300SURFST_T_F" Or Lictype = "86307VRDDES_T_F" Or _
                Lictype = "MKSIntegrityManager-Float" Or Lictype = "MKSSourceIntegrity-Float" Or Lictype = "DOORS" Or Lictype = "ClearQuest" Or Lictype = "ClearCase" Or _
                Lictype = "Allegro_Auth_HighSpeed_Option" Or Lictype = "Allegro_Design_Publisher" Or Lictype = "Concept_HDL_studio" Or Lictype = "PCB_design_studio" Or Lictype = "PCB_librarian_expert" Or Lictype = "PSpiceStudio" Or _
                Lictype = "MATLAB" Or Lictype = "SIMULINK" Or Lictype = "Video_and_Image_Blockset" Or Lictype = "Signal_Blocks" Or Lictype = "RTW_Embedded_Coder" Or Lictype = "Fixed_Point_Toolbox" Or _
                Lictype = "Image_Toolbox" Or Lictype = "MATLAB_Coder" Or Lictype = "MATLAB_Report_Gen" Or Lictype = "Signal_Toolbox" Or Lictype = "Real-Time_Workshop" Or Lictype = "SIMULINK_Report_Gen" Or _
                Lictype = "SL_Verification_Validation" Or Lictype = "Stateflow" Or Lictype = "Statistics_Toolbox" Or _
                Lictype = "77800MFS_T_F" Or Lictype = "86802PLC0000023_T_F" Or Lictype = "86387MFIP_T_F" Or Lictype = "85816MFAA_2012_0F" Or Lictype = "85819MFAM_2012_0F" Or Lictype = "86751MFIB_2017_0F" Or _
                Lictype = "86752MFIP_2017_0F" Or Lictype = "86555MFIP_2016_0F" Or Lictype = "86323MFIP_2015_0F" Or Lictype = "86155MFIP_2014_0F" Or Lictype = "86802PLC0000023_2017_0F" Or Lictype = "86754MFS_2017_0F" Or Lictype = "86558MFS_2016_0F" Or _
                Lictype = "86326MFS_2015_0F" Or Lictype = "86158MFS_2014_0F" Or Lictype = "87198MFIA_2019_0F" Or Lictype = "87016MFIA_2018_0F" Or Lictype = "86753MFIA_2017_0F" Or Lictype = "86557MFIA_2016_0F" Or Lictype = "86325MFIA_2015_0F" Or Lictype = "86156MFIA_2014_0F" Or _
                Lictype = "87196MFIP_2019_0F" Or Lictype = "87014MFIP_2018_0F" Or Lictype = "87032PLC0000023_2018_0F" Or Lictype = "87194MFS_2019_0F" Or Lictype = "87017MFS_2018_0F" Or _
                Lictype = "csct" Then
                    Dim GetLicRow As DataRow()
                    GetLicRow = Lic_InfoDT.Select("LIC_VER='" & LicVer & "' AND LIC_TYPE='" & Lictype & "' AND LIC_SRV='" & LicSrv & "'")
                    Try
                        If LicOnline <> 0 Then
                            For index1 As Integer = 0 To LicOnline - 1
                                CadType = GetLicRow(0).Item("CAD_TYPE")
                                LicSrv = GetLicRow(0).Item("LIC_SRV")
                                LicNa = GetLicRow(0).Item("LIC_NATION")
                                LicPro = GetLicRow(0).Item("LIC_PROPERTY")

                                GetLicRow(0).Item("LIC_NUM") = LicNum
                                GetLicRow(0).Item("LIC_ONLINE") = LicOnline
                                GetLicRow(0).Item("LIC_OFFLINE") = LicOffline

                                If strCAD = "COMMON" Then
                                    arrValue = SysOp.GetValue(arrNXLicInfo.Item(index + (index1 + 5)).ToString.Split(" "))
                                Else
                                    arrValue = SysOp.GetValue(arrNXLicInfo.Item(index + (index1 + 6)).ToString.Split(" "))
                                End If

                                If arrValue.Contains(strFilterforOffline) Then
                                    LicOffline = LicOffline + 1
                                End If

                                UserID = arrValue(0)

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
                                NewRow.Item("LIC_PROPERTY") = LicPro
                            Next
                        Else
                            CadType = GetLicRow(0).Item("CAD_TYPE")
                            LicSrv = GetLicRow(0).Item("LIC_SRV")
                            LicNa = GetLicRow(0).Item("LIC_NATION")
                            LicPro = GetLicRow(0).Item("LIC_PROPERTY")
                            GetLicRow(0).Item("LIC_NUM") = LicNum
                            GetLicRow(0).Item("LIC_ONLINE") = LicOnline
                            'GetLicRow(0).Item("LIC_OFFLINE") = LicOffline
                        End If
                    Catch ex As Exception
                        Console.WriteLine(Lictype & "/" & LicVer & "/" & " : Error!!!")
                        Console.WriteLine("ERROR :" & ex.Message)
                    End Try
                    GetLicRow(0).Item("LIC_OFFLINE") = LicOffline
                End If
            End If
        Next

        Return Lic_SrvDT


    End Function
    Private Function GetFLUENTLicenseServer_Information(ByVal Lic_InfoDT As DataTable, ByVal Lic_SrvDT As DataTable, ByVal arrNXLicInfo As ArrayList, ByVal strIP As String) As DataTable

        Dim arrLicSrvInfo As New ArrayList
        Dim OnlineUserNUM As Integer = 0
        Dim FilterStr As String = "Users of "
        Dim strFluentLic1 As String = "anshpc_pack"
        Dim strFluentLic2 As String = "acfd_2"
        For index As Integer = 12 To arrNXLicInfo.Count - 1
            Debug.Print(arrNXLicInfo.Item(index))
            If arrNXLicInfo.Item(index).ToString.Contains(FilterStr) Then
                If arrNXLicInfo.Item(index).ToString.Contains(strFluentLic1) Or arrNXLicInfo.Item(index).ToString.Contains(strFluentLic2) Then
                    Dim arrValue As New ArrayList
                    Dim CadType As String
                    Dim Lictype As String
                    Dim LicVer As String = "-"
                    Dim LicNum As Integer
                    Dim LicOnline As Integer
                    Dim LicOffline As Integer = 0
                    Dim UserID As String
                    Dim UserIP As String = "-"
                    Dim LicSrv As String = strIP
                    Dim LicNa As String
                    Dim LicDate As String = SysOp.SortDateTime(My.Computer.Clock.LocalTime)
                    Dim LicPro As String

                    arrValue = SysOp.GetValue(arrNXLicInfo.Item(index).ToString.Split(" "))
                    Lictype = arrValue(2).ToString.Replace(":", "")

                    Try
                        LicNum = arrValue(5)
                    Catch ex As Exception
                        LicNum = LicNum
                    End Try

                    Try
                        LicOnline = arrValue(10)
                    Catch ex As Exception
                        LicOnline = LicOnline
                    End Try

                    Dim GetLicRow As DataRow()
                    GetLicRow = Lic_InfoDT.Select("LIC_VER='" & LicVer & "' AND LIC_TYPE='" & Lictype & "' AND LIC_SRV='" & LicSrv & "'")
                    Try
                        If LicOnline <> 0 Then
                            Dim intIndex As Integer = index + 1
                            Do Until arrNXLicInfo.Item(intIndex).ToString.Contains(FilterStr)
                                Dim strUsedFilter As String = "start"
                                If arrNXLicInfo.Item(intIndex).ToString.Contains(strUsedFilter) Then
                                    CadType = GetLicRow(0).Item("CAD_TYPE")
                                    LicSrv = GetLicRow(0).Item("LIC_SRV")
                                    LicNa = GetLicRow(0).Item("LIC_NATION")
                                    LicPro = GetLicRow(0).Item("LIC_PROPERTY")

                                    GetLicRow(0).Item("LIC_NUM") = LicNum
                                    GetLicRow(0).Item("LIC_ONLINE") = LicOnline
                                    GetLicRow(0).Item("LIC_OFFLINE") = LicOffline

                                    arrValue = SysOp.GetValue(arrNXLicInfo.Item(intIndex).ToString.Split(" "))
                                    UserID = arrValue(0)

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
                                    NewRow.Item("LIC_PROPERTY") = LicPro
                                End If
                                intIndex = intIndex + 1
                            Loop
                        Else
                            CadType = GetLicRow(0).Item("CAD_TYPE")
                            LicSrv = GetLicRow(0).Item("LIC_SRV")
                            LicNa = GetLicRow(0).Item("LIC_NATION")
                            LicPro = GetLicRow(0).Item("LIC_PROPERTY")

                            GetLicRow(0).Item("LIC_NUM") = LicNum
                            GetLicRow(0).Item("LIC_ONLINE") = LicOnline
                            GetLicRow(0).Item("LIC_OFFLINE") = LicOffline


                        End If
                    Catch ex As Exception
                        Console.WriteLine(Lictype & "/" & LicVer & "/" & " : Error!!!")
                        Console.WriteLine("ERROR :" & ex.Message)
                    End Try
                End If
            End If
        Next

        Return Lic_SrvDT


    End Function
    Private Function GetHyperWorksLicenseServer_Information(ByVal Lic_InfoDT As DataTable, ByVal Lic_SrvDT As DataTable, ByVal arrNXLicInfo As ArrayList, ByVal strIP As String, ByVal strNa As String) As DataTable

        Dim arrLicSrvInfo As New ArrayList
        Dim OnlineUserNUM As Integer = 0
        Dim FilterStr As String = "Feature: "
        Dim arrHWLicList As New ArrayList
        Dim GetLicTypeRow As DataRow()
        Dim arrLicTypelist As New ArrayList
        GetLicTypeRow = Lic_InfoDT.Select("CAD_TYPE = 'HYPERWORKS'")
        For i As Integer = 0 To GetLicTypeRow.Count - 1
            arrLicTypelist.Add(GetLicTypeRow(i).Item("LIC_TYPE"))
        Next

        For index As Integer = 0 To arrNXLicInfo.Count - 1
            Debug.Print(arrNXLicInfo.Item(index))
            If arrNXLicInfo.Item(index).ToString.Contains(FilterStr) Then

                Dim arrValue As New ArrayList
                Dim strSubFilter As String
                Dim CadType As String
                Dim Lictype As String
                Dim LicVer As String
                Dim LicNum As Integer
                Dim LicOnline As Integer
                Dim LicOffline As Integer = 0
                Dim UserID As String
                Dim UserIP As String = "-"
                Dim LicSrv As String = strIP
                Dim LicNa As String
                Dim LicDate As String = SysOp.SortDateTime(My.Computer.Clock.LocalTime)
                Dim LicPro As String
                Dim LicToken As Integer = 0
                Dim arrLicTypeFilter As New ArrayList
                Dim arrLicTokenFilter As New ArrayList

                If strNa = "KOR" Then
                    LicVer = "19.0"
                Else
                    LicVer = "18.0"
                End If

                arrValue = SysOp.GetValue(arrNXLicInfo.Item(index).ToString.Split(" "))
                If arrLicTypelist.Contains(arrValue(1)) Then

                    Lictype = arrValue(1).ToString

                    arrValue = SysOp.GetValue(arrNXLicInfo.Item(index + 4).ToString.Split(" "))

                    Try
                        LicNum = arrValue(2)
                    Catch ex As Exception
                        LicNum = LicNum
                    End Try

                    Try
                        LicOnline = arrValue(0)
                    Catch ex As Exception
                        LicOnline = LicOnline
                    End Try

                    Dim GetLicRow As DataRow()
                    GetLicRow = Lic_InfoDT.Select("LIC_VER='" & LicVer & "' AND LIC_TYPE='" & Lictype & "' AND LIC_SRV='" & LicSrv & "'")
                    Try
                        If LicOnline <> 0 Then
                            Dim intIndex As Integer = index + 1
                            Dim strEnd As String = "denial(s)"
                            Do Until arrNXLicInfo.Item(intIndex).ToString.Contains(strEnd)
                                Dim strUsedFilter As String = "license(s) used by"
                                If arrNXLicInfo.Item(intIndex).ToString.Contains(strUsedFilter) Then
                                    arrValue = SysOp.GetValue(arrNXLicInfo.Item(intIndex).ToString.Split(" "))
                                    LicToken = CInt(arrValue(0).ToString)
                                    CadType = GetLicRow(0).Item("CAD_TYPE")
                                    LicSrv = GetLicRow(0).Item("LIC_SRV")
                                    LicNa = GetLicRow(0).Item("LIC_NATION")
                                    LicPro = GetLicRow(0).Item("LIC_PROPERTY")

                                    GetLicRow(0).Item("LIC_NUM") = LicNum
                                    GetLicRow(0).Item("LIC_ONLINE") = LicOnline
                                    GetLicRow(0).Item("LIC_OFFLINE") = LicOffline

                                    strSubFilter = arrValue(5).ToString
                                    strSubFilter = strSubFilter.Remove(strSubFilter.Length - 1)
                                    strSubFilter = strSubFilter.Substring(1)
                                    arrValue = SysOp.GetValue(arrValue.Item(4).ToString.Split("@"))
                                    UserID = arrValue(0)
                                    UserIP = strSubFilter

                                    If Lictype = "HyperWorks" Then
                                        'If Lictype = "HyperWorks" Or Lictype = "GlobalZoneAP" Then
                                        If LicToken > 6000 Then
                                            If arrLicTypeFilter.Contains(strSubFilter) = False Then

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
                                                NewRow.Item("LIC_PROPERTY") = LicPro
                                                NewRow.Item("TOKEN_USED") = LicToken
                                                arrLicTypeFilter.Add(UserIP)
                                                Debug.Print(Lictype & "," & UserID & "," & UserIP & "," & LicToken.ToString)
                                            End If
                                        End If
                                        'Else
                                        '    If Lictype = "HWHyperGraph" Then
                                        '        If LicToken >= 6000 Then
                                        '            Dim strUniqueID As String
                                        '            Dim arrUnique As New ArrayList
                                        '            arrUnique = SysOp.GetValue(arrNXLicInfo.Item(intIndex + 2).ToString.Split(" "))
                                        '            strUniqueID = arrUnique.Item(4).ToString
                                        '            If arrLicTokenFilter.Contains(strUniqueID) = False Then
                                        '                Dim NewRow As DataRow
                                        '                NewRow = Lic_SrvDT.Rows.Add()
                                        '                NewRow.Item("CAD_TYPE") = CadType
                                        '                NewRow.Item("LIC_TYPE") = Lictype
                                        '                NewRow.Item("LIC_VER") = LicVer
                                        '                NewRow.Item("LIC_SRV") = LicSrv
                                        '                NewRow.Item("USER_ID") = UserID
                                        '                NewRow.Item("IP_INFO") = UserIP
                                        '                NewRow.Item("LIC_NATION") = LicNa
                                        '                NewRow.Item("LIC_DATE") = LicDate
                                        '                NewRow.Item("LIC_PROPERTY") = LicPro
                                        '                NewRow.Item("TOKEN_USED") = LicToken
                                        '                arrLicTokenFilter.Add(strUniqueID)
                                        '                Debug.Print(Lictype & "," & UserID & "," & UserIP & "," & LicToken.ToString)
                                        '            End If
                                        '        End If
                                        '    ElseIf Lictype = "HWPDataManager" Then
                                        '        If LicToken >= 2000 Then
                                        '            Dim strUniqueID As String
                                        '            Dim arrUnique As New ArrayList
                                        '            arrUnique = SysOp.GetValue(arrNXLicInfo.Item(intIndex + 2).ToString.Split(" "))
                                        '            strUniqueID = arrUnique.Item(4).ToString
                                        '            If arrLicTokenFilter.Contains(strUniqueID) = False Then
                                        '                Dim NewRow As DataRow
                                        '                NewRow = Lic_SrvDT.Rows.Add()
                                        '                NewRow.Item("CAD_TYPE") = CadType
                                        '                NewRow.Item("LIC_TYPE") = Lictype
                                        '                NewRow.Item("LIC_VER") = LicVer
                                        '                NewRow.Item("LIC_SRV") = LicSrv
                                        '                NewRow.Item("USER_ID") = UserID
                                        '                NewRow.Item("IP_INFO") = UserIP
                                        '                NewRow.Item("LIC_NATION") = LicNa
                                        '                NewRow.Item("LIC_DATE") = LicDate
                                        '                NewRow.Item("LIC_PROPERTY") = LicPro
                                        '                NewRow.Item("TOKEN_USED") = LicToken
                                        '                arrLicTokenFilter.Add(strUniqueID)
                                        '                Debug.Print(Lictype & "," & UserID & "," & UserIP & "," & LicToken.ToString)
                                        '            End If
                                        '        End If
                                        '    Else
                                        '        If LicToken > 6000 Then
                                        '            Dim strUniqueID As String
                                        '            Dim arrUnique As New ArrayList
                                        '            arrUnique = SysOp.GetValue(arrNXLicInfo.Item(intIndex + 2).ToString.Split(" "))
                                        '            strUniqueID = arrUnique.Item(4).ToString
                                        '            If arrLicTokenFilter.Contains(strUniqueID) = False Then
                                        '                Dim NewRow As DataRow
                                        '                NewRow = Lic_SrvDT.Rows.Add()
                                        '                NewRow.Item("CAD_TYPE") = CadType
                                        '                NewRow.Item("LIC_TYPE") = Lictype
                                        '                NewRow.Item("LIC_VER") = LicVer
                                        '                NewRow.Item("LIC_SRV") = LicSrv
                                        '                NewRow.Item("USER_ID") = UserID
                                        '                NewRow.Item("IP_INFO") = UserIP
                                        '                NewRow.Item("LIC_NATION") = LicNa
                                        '                NewRow.Item("LIC_DATE") = LicDate
                                        '                NewRow.Item("LIC_PROPERTY") = LicPro
                                        '                NewRow.Item("TOKEN_USED") = LicToken
                                        '                arrLicTokenFilter.Add(strUniqueID)
                                        '                Debug.Print(Lictype & "," & UserID & "," & UserIP & "," & LicToken.ToString)
                                        '            End If
                                        '        End If
                                        '    End If
                                    End If
                                End If
                                intIndex = intIndex + 1
                            Loop
                        Else
                            CadType = GetLicRow(0).Item("CAD_TYPE")
                            LicSrv = GetLicRow(0).Item("LIC_SRV")
                            LicNa = GetLicRow(0).Item("LIC_NATION")
                            LicPro = GetLicRow(0).Item("LIC_PROPERTY")

                            GetLicRow(0).Item("LIC_NUM") = LicNum
                            GetLicRow(0).Item("LIC_ONLINE") = LicOnline
                            GetLicRow(0).Item("LIC_OFFLINE") = LicOffline
                        End If
                    Catch ex As Exception
                        Console.WriteLine(Lictype & "/" & LicVer & "/" & " : Error!!!")
                        Console.WriteLine("ERROR :" & ex.Message)
                    End Try

                End If
            End If
        Next

        Return Lic_SrvDT
    End Function
    Private Function GetMOLDFLOW_Information(ByVal Lic_InfoDT As DataTable, ByVal Lic_SrvDT As DataTable, ByVal arrNXLicInfo As ArrayList, ByVal strIP As String, ByVal strCAD As String) As DataTable

        Dim iNX13100N As Integer = 0
        Dim arrLicSrvInfo As New ArrayList
        Dim OnlineUserNUM As Integer = 0
        Dim FilterStr As String = "Users of "
        For index As Integer = 11 To arrNXLicInfo.Count - 1
            Debug.Print(arrNXLicInfo.Item(index))
            If arrNXLicInfo.Item(index).ToString.Contains(FilterStr) Then
                Dim arrValue As New ArrayList
                Dim CadType As String
                Dim Lictype As String
                Dim LicVer As String = "-"
                Dim LicNum As Integer
                Dim LicOnline As Integer
                Dim LicOffline As Integer = 0
                Dim UserID As String
                Dim UserIP As String = "-"
                Dim LicSrv As String = strIP
                Dim LicNa As String
                Dim LicDate As String = SysOp.SortDateTime(My.Computer.Clock.LocalTime)
                Dim LicPro As String

                arrValue = SysOp.GetValue(arrNXLicInfo.Item(index).ToString.Split(" "))
                Lictype = arrValue(2).ToString.Replace(":", "")

                Try
                    LicNum = arrValue(5)
                Catch ex As Exception
                    LicNum = LicNum
                End Try

                Try
                    LicOnline = arrValue(10)
                Catch ex As Exception
                    LicOnline = LicOnline
                End Try

                'If LicOnline <> 0 Then
                '    arrValue = SysOp.GetValue(arrNXLicInfo.Item(index + 2).ToString.Split(" "))
                '    LicVer = arrValue(1).ToString.Replace(",", "")
                'Else
                '    'LMTOOLS 는 사용자가 없을 경우 버전 및 전체 사용자가 나오지 않아서 사용자가 없는 경우를 셋팅해줘야함
                '    Select Case Lictype
                '        Case "77800MFS_T_F"
                '            LicVer = "v1.000"
                '        Case "86387MFIP_T_F"
                '            LicVer = "v1.000"
                '        Case "77400MFIA_T_F"
                '            LicVer = "v1.000"
                '    End Select
                'End If


                If Lictype = "77800MFS_T_F" Or Lictype = "86387MFIP_T_F" Or Lictype = "77400MFIA_T_F" Then
                    Dim GetLicRow As DataRow()
                    GetLicRow = Lic_InfoDT.Select("LIC_VER='" & LicVer & "' AND LIC_TYPE='" & Lictype & "' AND LIC_SRV='" & LicSrv & "'")
                    Try
                        If LicOnline <> 0 Then
                            Dim intIndex As Integer = index + 6
                            Dim intOnlineCount As Integer = 0
                            Do Until arrNXLicInfo.Item(intIndex).ToString.Contains(FilterStr)
                                If arrNXLicInfo.Item(intIndex).ToString <> "" Then
                                    If arrNXLicInfo.Item(intIndex).ToString.Contains(Lictype.ToString) = False Then

                                        CadType = GetLicRow(0).Item("CAD_TYPE")
                                        LicSrv = GetLicRow(0).Item("LIC_SRV")
                                        LicNa = GetLicRow(0).Item("LIC_NATION")
                                        LicPro = GetLicRow(0).Item("LIC_PROPERTY")

                                        GetLicRow(0).Item("LIC_NUM") = LicNum
                                        GetLicRow(0).Item("LIC_ONLINE") = LicOnline
                                        GetLicRow(0).Item("LIC_OFFLINE") = LicOffline

                                        arrValue = SysOp.GetValue(arrNXLicInfo.Item(index + (intOnlineCount + 6)).ToString.Split(" "))

                                        UserID = arrValue(0)

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
                                        NewRow.Item("LIC_PROPERTY") = LicPro

                                        intOnlineCount = intOnlineCount + 1

                                    Else
                                        Dim intCountforPosition As Integer
                                        intCountforPosition = LicOnline - intOnlineCount

                                        For i As Integer = 0 To (LicOnline - intOnlineCount) - 1

                                            CadType = GetLicRow(0).Item("CAD_TYPE")
                                            LicSrv = GetLicRow(0).Item("LIC_SRV")
                                            LicNa = GetLicRow(0).Item("LIC_NATION")
                                            LicPro = GetLicRow(0).Item("LIC_PROPERTY")

                                            GetLicRow(0).Item("LIC_NUM") = LicNum
                                            GetLicRow(0).Item("LIC_ONLINE") = LicOnline
                                            GetLicRow(0).Item("LIC_OFFLINE") = LicOffline

                                            arrValue = SysOp.GetValue(arrNXLicInfo.Item(index + (intOnlineCount + 11)).ToString.Split(" "))

                                            UserID = arrValue(0)

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
                                            NewRow.Item("LIC_PROPERTY") = LicPro

                                            intOnlineCount = intOnlineCount + 1

                                        Next
                                        'intIndex = intIndex + 5
                                        intIndex = intIndex + (intCountforPosition + 4)
                                    End If
                                End If

                                intIndex = intIndex + 1

                            Loop
                        Else
                            CadType = GetLicRow(0).Item("CAD_TYPE")
                            LicSrv = GetLicRow(0).Item("LIC_SRV")
                            LicNa = GetLicRow(0).Item("LIC_NATION")
                            LicPro = GetLicRow(0).Item("LIC_PROPERTY")

                            GetLicRow(0).Item("LIC_NUM") = LicNum
                            GetLicRow(0).Item("LIC_ONLINE") = LicOnline
                            GetLicRow(0).Item("LIC_OFFLINE") = LicOffline
                        End If
                    Catch ex As Exception
                        Console.WriteLine(Lictype & "/" & LicVer & "/" & " : Error!!!")
                        Console.WriteLine("ERROR :" & ex.Message)
                    End Try
                End If
            End If
        Next

        Return Lic_SrvDT

    End Function
    Private Function GetABAQUSLicenseServer_Information(ByVal Lic_InfoDT As DataTable, ByVal Lic_SrvDT As DataTable, ByVal arrNXLicInfo As ArrayList, ByVal strIP As String) As DataTable

        Dim arrLicSrvInfo As New ArrayList
        Dim OnlineUserNUM As Integer = 0
        Dim strFilter As String = "Users of abaqus"
        Dim strFilterforDoLoop As String = "Users of"

        For index As Integer = 12 To arrNXLicInfo.Count - 1
            Debug.Print(arrNXLicInfo.Item(index))
            If arrNXLicInfo.Item(index).ToString.Contains(strFilter) Then

                Dim arrValue As New ArrayList
                Dim CadType As String
                Dim Lictype As String
                Dim LicVer As String = "v62.0"
                Dim LicNum As Integer
                Dim LicOnline As Integer
                Dim LicOffline As Integer = 0
                Dim UserID As String
                Dim UserIP As String = "-"
                Dim LicSrv As String = strIP
                Dim LicNa As String
                Dim LicDate As String = SysOp.SortDateTime(My.Computer.Clock.LocalTime)
                Dim LicPro As String

                arrValue = SysOp.GetValue(arrNXLicInfo.Item(index).ToString.Split(" "))
                Lictype = arrValue(2).ToString.Replace(":", "")

                Try
                    LicNum = arrValue(5)
                Catch ex As Exception
                    LicNum = LicNum
                End Try

                Try
                    LicOnline = arrValue(10)
                Catch ex As Exception
                    LicOnline = LicOnline
                End Try

                Dim GetLicRow As DataRow()
                GetLicRow = Lic_InfoDT.Select("LIC_VER='" & LicVer & "' AND LIC_TYPE='" & Lictype & "' AND LIC_SRV='" & LicSrv & "'")
                Try
                    If LicOnline <> 0 Then
                        Dim intIndex As Integer = index + 1
                        Do Until arrNXLicInfo.Item(intIndex).ToString.Contains(strFilterforDoLoop)
                            Dim strUsedFilter As String = "start"
                            If arrNXLicInfo.Item(intIndex).ToString.Contains(strUsedFilter) Then
                                CadType = GetLicRow(0).Item("CAD_TYPE")
                                LicSrv = GetLicRow(0).Item("LIC_SRV")
                                LicNa = GetLicRow(0).Item("LIC_NATION")
                                LicPro = GetLicRow(0).Item("LIC_PROPERTY")

                                GetLicRow(0).Item("LIC_NUM") = LicNum
                                GetLicRow(0).Item("LIC_ONLINE") = LicOnline
                                GetLicRow(0).Item("LIC_OFFLINE") = LicOffline

                                arrValue = SysOp.GetValue(arrNXLicInfo.Item(intIndex).ToString.Split(" "))
                                UserID = arrValue(0)

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
                                NewRow.Item("LIC_PROPERTY") = LicPro
                            End If
                            intIndex = intIndex + 1
                        Loop
                    Else
                        CadType = GetLicRow(0).Item("CAD_TYPE")
                        LicSrv = GetLicRow(0).Item("LIC_SRV")
                        LicNa = GetLicRow(0).Item("LIC_NATION")
                        LicPro = GetLicRow(0).Item("LIC_PROPERTY")

                        GetLicRow(0).Item("LIC_NUM") = LicNum
                        GetLicRow(0).Item("LIC_ONLINE") = LicOnline
                        GetLicRow(0).Item("LIC_OFFLINE") = LicOffline


                    End If
                Catch ex As Exception
                    Console.WriteLine(Lictype & "/" & LicVer & "/" & " : Error!!!")
                    Console.WriteLine("ERROR :" & ex.Message)
                End Try
            End If
        Next

        Return Lic_SrvDT


    End Function
    Private Function GetFlexLm_DoLoop_LicenseServer_Information(ByVal Lic_InfoDT As DataTable, ByVal Lic_SrvDT As DataTable, ByVal arrNXLicInfo As ArrayList, ByVal strIP As String, ByVal strCadType As String) As DataTable
        Dim arrLicSrvInfo As New ArrayList
        Dim OnlineUserNUM As Integer = 0
        Dim FilterStr As String = "Users of "
        For index As Integer = 11 To arrNXLicInfo.Count - 1
            Debug.Print(arrNXLicInfo.Item(index))
            If arrNXLicInfo.Item(index).ToString.Contains(FilterStr) Then
                Dim arrValue As New ArrayList
                Dim CadType As String
                Dim Lictype As String
                Dim LicVer As String = "-"
                Dim LicNum As Integer
                Dim LicOnline As Integer
                Dim LicOffline As Integer = 0
                Dim UserID As String
                Dim UserIP As String = "-"
                Dim LicSrv As String = strIP
                Dim LicNa As String
                Dim LicDate As String = SysOp.SortDateTime(My.Computer.Clock.LocalTime)
                Dim LicPro As String
                Dim bolLictype As Boolean = False


                arrValue = SysOp.GetValue(arrNXLicInfo.Item(index).ToString.Split(" "))
                Lictype = arrValue(2).ToString.Replace(":", "")

                Try
                    LicNum = arrValue(5)
                Catch ex As Exception
                    LicNum = LicNum
                End Try

                Try
                    LicOnline = arrValue(10)
                Catch ex As Exception
                    LicOnline = LicOnline
                End Try


                'ANSYS/FLUENT 라이선스 항목 중첩 => SW 구분해서 모니터링하기 위해 구분
                If strCadType = "ANSYS" Then
                    '    If Lictype = "q3d_desktop" Or Lictype = "si2d_gui" Or Lictype = "si2d_solve" Or Lictype = "si3d_gui" Or Lictype = "si3d_solve" Or Lictype = "simplorer_CProgrInterface" Or Lictype = "simplorer_gui" Or Lictype = "simplorer_LibSMPS" Or _
                    'Lictype = "simplorer_vhdlams" Or Lictype = "simplorer_control" Or Lictype = "simplorer_modelica" Or Lictype = "electronics_desktop" Or Lictype = "simplorer_sim_entry" Or Lictype = "hfsshpc_pack" Or Lictype = "hfssie_solve" Or Lictype = "hfssie_gui" Or _
                    'Lictype = "fwspice_export" Or Lictype = "al4allegro" Or Lictype = "al4ansoft" Or Lictype = "al4apd" Or Lictype = "al4boardstation" Or Lictype = "al4encore" Or Lictype = "al4expedition" Or Lictype = "al4generic" Or _
                    'Lictype = "al4powerpcb" Or Lictype = "al4virtuoso" Or Lictype = "al4zuken" Or Lictype = "alinks_gui" Or Lictype = "al4cadvance" Or Lictype = "al4cds" Or Lictype = "al4gem" Or Lictype = "al4first" Or _
                    'Lictype = "al4odb++" Or Lictype = "xlate_catia4" Or Lictype = "xlate_catia5" Or Lictype = "xlate_parasolid" Or Lictype = "xlate_unigraphics" Or Lictype = "designer_desktop" Or Lictype = "ensemble_gui" Or Lictype = "nexxim_gui" Or _
                    'Lictype = "nexxim_netlist" Or Lictype = "serenade_gui" Or Lictype = "symphony_gui" Or Lictype = "nexxim_dc" Or Lictype = "nexxim_eye" Or Lictype = "nexxim_tran" Or Lictype = "filter_synthesis" Or Lictype = "serenade_adv_sim" Or _
                    'Lictype = "serenade_linear" Or Lictype = "nexxim_icda" Or Lictype = "nexxim_ami" Or Lictype = "hfss_solve" Or Lictype = "emit_legacy_gui" Or Lictype = "savant_legacy_gui" Or Lictype = "siwave_gui" Or Lictype = "designer_hspice" Or _
                    'Lictype = "siwave_level1" Or Lictype = "siwave_level2" Or Lictype = "siwave_level3" Or Lictype = "simplorer_desktop" Or Lictype = "simplorer_sim" Or Lictype = "piproe" Or Lictype = "agppi" Or Lictype = "aice_mesher" Or _
                    'Lictype = "aice_pak" Or Lictype = "aice_solv" Or Lictype = "aiiges" Or Lictype = "electronics2d_gui" Or Lictype = "electronics3d_gui" Or Lictype = "electronicsckt_gui" Or Lictype = "dsdxm" Or Lictype = "a_spaceclaim_dirmod" Or _
                    'Lictype = "anshpc_pack" Or Lictype = "m2dfs_qs_solve" Or Lictype = "m2dfs_solve" Or Lictype = "hfss_transient_solve" Or Lictype = "ansoft_distrib_engine" Or Lictype = "simplorer_twin_models" Then
                    '        bolLictype = True
                    '    End If
                    If Lictype = "electronics_desktop" Or Lictype = "siwave_gui" Or Lictype = "simplorer_desktop" Or Lictype = "alinks_gui" Then
                        bolLictype = True
                    End If
                ElseIf strCadType = "FLUENT" Then
                    If Lictype = "anshpc_pack" Or Lictype = "electronics_desktop" Or Lictype = "hfss_solve" Or Lictype = "siwave_gui" Or Lictype = "cfd_solve_level1" Then
                        bolLictype = True
                    End If
                ElseIf strCadType = "CRADLE" Then
                    If Lictype = "SCTWPP" Or Lictype = "SCTMPIJOB" Then
                        bolLictype = True
                    End If
                Else
                    bolLictype = False
                End If

                '1~10열 ANSYS, 11열 FLUENT 12열 CRADLE
                'If Lictype = "q3d_desktop" Or Lictype = "si2d_gui" Or Lictype = "si2d_solve" Or Lictype = "si3d_gui" Or Lictype = "si3d_solve" Or Lictype = "simplorer_CProgrInterface" Or Lictype = "simplorer_gui" Or Lictype = "simplorer_LibSMPS" Or _
                'Lictype = "simplorer_vhdlams" Or Lictype = "simplorer_control" Or Lictype = "simplorer_modelica" Or Lictype = "electronics_desktop" Or Lictype = "simplorer_sim_entry" Or Lictype = "hfsshpc_pack" Or Lictype = "hfssie_solve" Or Lictype = "hfssie_gui" Or _
                'Lictype = "fwspice_export" Or Lictype = "al4allegro" Or Lictype = "al4ansoft" Or Lictype = "al4apd" Or Lictype = "al4boardstation" Or Lictype = "al4encore" Or Lictype = "al4expedition" Or Lictype = "al4generic" Or _
                'Lictype = "al4powerpcb" Or Lictype = "al4virtuoso" Or Lictype = "al4zuken" Or Lictype = "alinks_gui" Or Lictype = "al4cadvance" Or Lictype = "al4cds" Or Lictype = "al4gem" Or Lictype = "al4first" Or _
                'Lictype = "al4odb++" Or Lictype = "xlate_catia4" Or Lictype = "xlate_catia5" Or Lictype = "xlate_parasolid" Or Lictype = "xlate_unigraphics" Or Lictype = "designer_desktop" Or Lictype = "ensemble_gui" Or Lictype = "nexxim_gui" Or _
                'Lictype = "nexxim_netlist" Or Lictype = "serenade_gui" Or Lictype = "symphony_gui" Or Lictype = "nexxim_dc" Or Lictype = "nexxim_eye" Or Lictype = "nexxim_tran" Or Lictype = "filter_synthesis" Or Lictype = "serenade_adv_sim" Or _
                'Lictype = "serenade_linear" Or Lictype = "nexxim_icda" Or Lictype = "nexxim_ami" Or Lictype = "hfss_solve" Or Lictype = "emit_legacy_gui" Or Lictype = "savant_legacy_gui" Or Lictype = "siwave_gui" Or Lictype = "designer_hspice" Or _
                'Lictype = "siwave_level1" Or Lictype = "siwave_level2" Or Lictype = "siwave_level3" Or Lictype = "simplorer_desktop" Or Lictype = "simplorer_sim" Or Lictype = "piproe" Or Lictype = "agppi" Or Lictype = "aice_mesher" Or _
                'Lictype = "aice_pak" Or Lictype = "aice_solv" Or Lictype = "aiiges" Or Lictype = "electronics2d_gui" Or Lictype = "electronics3d_gui" Or Lictype = "electronicsckt_gui" Or Lictype = "dsdxm" Or Lictype = "a_spaceclaim_dirmod" Or _
                'Lictype = "anshpc_pack" Or Lictype = "m2dfs_qs_solve" Or Lictype = "m2dfs_solve" Or Lictype = "hfss_transient_solve" Or Lictype = "ansoft_distrib_engine" Or Lictype = "simplorer_twin_models" Or _
                'Lictype = "cfd_base" Or _
                'Lictype = "SCTWPP" Or Lictype = "SCTOTHER" Or Lictype = "SCOTHER" Or Lictype = "SCTMPIJOB" Or Lictype = "SCTSOLMPI" Or Lictype = "SCTMPIGRP" Or Lictype = "SCTOTHERMPI" Then

                If bolLictype = True Then
                    Dim GetLicRow As DataRow()
                    GetLicRow = Lic_InfoDT.Select("LIC_VER='" & LicVer & "' AND LIC_TYPE='" & Lictype & "' AND LIC_SRV='" & LicSrv & "'")
                    Try
                        If LicOnline <> 0 Then
                            Dim intIndexforExtract As Integer
                            intIndexforExtract = index
                            Do Until arrNXLicInfo.Item(intIndexforExtract + 6).ToString.Contains("Users of")
                                Debug.Print(arrNXLicInfo.Item(intIndexforExtract + 6).ToString)
                                If arrNXLicInfo.Item(intIndexforExtract + 6).ToString <> "" Then
                                    If arrNXLicInfo.Item(intIndexforExtract + 6).ToString.Contains(Lictype) = False Then
                                        If arrNXLicInfo.Item(intIndexforExtract + 6).ToString.Contains("vendor_string") = False Then
                                            If arrNXLicInfo.Item(intIndexforExtract + 6).ToString.Contains("floating") = False Then
                                                CadType = GetLicRow(0).Item("CAD_TYPE")
                                                LicSrv = GetLicRow(0).Item("LIC_SRV")
                                                LicNa = GetLicRow(0).Item("LIC_NATION")
                                                LicPro = GetLicRow(0).Item("LIC_PROPERTY")
                                                Try
                                                    GetLicRow(0).Item("LIC_NUM") = LicNum
                                                    GetLicRow(0).Item("LIC_ONLINE") = LicOnline
                                                    GetLicRow(0).Item("LIC_OFFLINE") = LicOffline

                                                    arrValue = SysOp.GetValue(arrNXLicInfo.Item(intIndexforExtract + 6).ToString.Split(" "))

                                                    'If strCAD = "COMMON" Then
                                                    '    arrValue = SysOp.GetValue(arrNXLicInfo.Item(index + (index1 + 5)).ToString.Split(" "))
                                                    'Else
                                                    '    arrValue = SysOp.GetValue(arrNXLicInfo.Item(index + (index1 + 6)).ToString.Split(" "))
                                                    'End If

                                                    UserID = arrValue(0)

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
                                                    NewRow.Item("LIC_PROPERTY") = LicPro
                                                Catch ex As Exception
                                                    Debug.Print(ex.ToString)
                                                End Try
                                            End If
                                        End If
                                    End If
                                End If
                                intIndexforExtract = intIndexforExtract + 1
                            Loop


                        Else
                            CadType = GetLicRow(0).Item("CAD_TYPE")
                            LicSrv = GetLicRow(0).Item("LIC_SRV")
                            LicNa = GetLicRow(0).Item("LIC_NATION")
                            LicPro = GetLicRow(0).Item("LIC_PROPERTY")

                            GetLicRow(0).Item("LIC_NUM") = LicNum
                            GetLicRow(0).Item("LIC_ONLINE") = LicOnline
                            GetLicRow(0).Item("LIC_OFFLINE") = LicOffline
                        End If
                    Catch ex As Exception
                        Console.WriteLine(Lictype & "/" & LicVer & "/" & " : Error!!!")
                        Console.WriteLine("ERROR :" & ex.Message)
                    End Try
                End If
            End If
        Next

        Return Lic_SrvDT

    End Function
    Private Function GetFlexLm_NoVersion_LicenseServer_Information(ByVal Lic_InfoDT As DataTable, ByVal Lic_SrvDT As DataTable, ByVal arrNXLicInfo As ArrayList, ByVal strIP As String, ByVal strCAD As String) As DataTable

        Dim arrLicSrvInfo As New ArrayList
        Dim FilterStr As String = "Users of "
        For index As Integer = 11 To arrNXLicInfo.Count - 1
            Debug.Print(arrNXLicInfo.Item(index))
            If arrNXLicInfo.Item(index).ToString.Contains(FilterStr) Then
                Dim arrValue As New ArrayList
                Dim CadType As String
                Dim Lictype As String
                Dim LicVer As String = "-"
                Dim LicNum As Integer
                Dim LicOnline As Integer
                Dim LicOffline As Integer = 0
                Dim UserID As String
                Dim UserIP As String = "-"
                Dim LicSrv As String = strIP
                Dim LicNa As String
                Dim LicDate As String = SysOp.SortDateTime(My.Computer.Clock.LocalTime)
                Dim LicPro As String

                arrValue = SysOp.GetValue(arrNXLicInfo.Item(index).ToString.Split(" "))
                Lictype = arrValue(2).ToString.Replace(":", "")

                Try
                    LicNum = arrValue(5)
                Catch ex As Exception
                    LicNum = LicNum
                End Try

                Try
                    LicOnline = arrValue(10)
                Catch ex As Exception
                    LicOnline = LicOnline
                End Try


                'DFPRE_APPCORE:DAFUL ccmppower:STARCCM primer:OASYS SherlockClient:SHERLOCK start:CST
                If Lictype = "DFPRE_APPCORE" Or Lictype = "ccmppower" Or Lictype = "primer" Or Lictype = "SherlockClient" Or Lictype = "start" Then
                    Dim GetLicRow As DataRow()
                    GetLicRow = Lic_InfoDT.Select("LIC_VER='" & LicVer & "' AND LIC_TYPE='" & Lictype & "' AND LIC_SRV='" & LicSrv & "'")
                    Try
                        If LicOnline <> 0 Then
                            For index1 As Integer = 0 To LicOnline - 1
                                CadType = GetLicRow(0).Item("CAD_TYPE")
                                LicSrv = GetLicRow(0).Item("LIC_SRV")
                                LicNa = GetLicRow(0).Item("LIC_NATION")
                                LicPro = GetLicRow(0).Item("LIC_PROPERTY")

                                GetLicRow(0).Item("LIC_NUM") = LicNum
                                GetLicRow(0).Item("LIC_ONLINE") = LicOnline
                                GetLicRow(0).Item("LIC_OFFLINE") = LicOffline

                                If strCAD = "COMMON" Then
                                    arrValue = SysOp.GetValue(arrNXLicInfo.Item(index + (index1 + 5)).ToString.Split(" "))
                                Else
                                    arrValue = SysOp.GetValue(arrNXLicInfo.Item(index + (index1 + 6)).ToString.Split(" "))
                                End If

                                UserID = arrValue(0)

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
                                NewRow.Item("LIC_PROPERTY") = LicPro
                            Next
                        Else
                            CadType = GetLicRow(0).Item("CAD_TYPE")
                            LicSrv = GetLicRow(0).Item("LIC_SRV")
                            LicNa = GetLicRow(0).Item("LIC_NATION")
                            LicPro = GetLicRow(0).Item("LIC_PROPERTY")

                            GetLicRow(0).Item("LIC_NUM") = LicNum
                            GetLicRow(0).Item("LIC_ONLINE") = LicOnline
                            GetLicRow(0).Item("LIC_OFFLINE") = LicOffline
                        End If
                    Catch ex As Exception
                        Console.WriteLine(Lictype & "/" & LicVer & "/" & " : Error!!!")
                        Console.WriteLine("ERROR :" & ex.Message)
                    End Try
                End If
            End If
        Next

        Return Lic_SrvDT

    End Function
End Class
