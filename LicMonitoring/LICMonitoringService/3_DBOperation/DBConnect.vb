Imports System.Data.SqlClient
Imports System.Data.OracleClient


'=======================테이블 정보(중요,StartMonitoring 에서 테이블 변경시 반드시 변경해줄것)======================
'Lic_SrvDT.Columns.Add("CAD_TYPE")        '0
'Lic_SrvDT.Columns.Add("LIC_TYPE")        '1
'Lic_SrvDT.Columns.Add("LIC_VER")         '2
'Lic_SrvDT.Columns.Add("LIC_SRV")         '3
'Lic_SrvDT.Columns.Add("USER_ID")         '4
'Lic_SrvDT.Columns.Add("NAME")            '5
'Lic_SrvDT.Columns.Add("DEPT_ID")         '6
'Lic_SrvDT.Columns.Add("DEPT_NAME")       '7
'Lic_SrvDT.Columns.Add("IP_INFO")         '8
'Lic_SrvDT.Columns.Add("BIZ_CD")          '9
'Lic_SrvDT.Columns.Add("M_ORG_NM")        '10
'Lic_SrvDT.Columns.Add("C_ORG_CD")        '11
'Lic_SrvDT.Columns.Add("C_DRG_NM")        '12
'Lic_SrvDT.Columns.Add("LIC_NATION")      '13
'====================================================================================================================

Public Class DBConnect

    Dim oSysOp As New SysOperation

    Public Function _CAD_SQLConnect() As SqlConnection

        Dim SQLConn As New SqlConnection
        Dim strCategory As String = "DB(CAD) Connection"
        SQLConn.ConnectionString = "Data Source=172.100.100.85;Initial Catalog=SLCAD;User ID=sa;Password=Samlip50;Connect Timeout=60000"

        Try
            SQLConn.Open()
            Debug.Print("CAD DB Connection")
            Return SQLConn
        Catch ex As Exception
            Console.WriteLine("SQLConnect ERROR : " & ex.Message)
            oSysOp.SendMail_forError(strCategory, "DB Connection Error" & ex.Message)
            Return SQLConn
            End
        End Try

    End Function

    Public Function _EP_Connect() As OracleConnection

        Dim ORAConn As New OracleConnection
        Dim strConnection As String
        strConnection = "data source=210.105.188.8:1521/SLEP;User Id=fastuser;Password=fast000"
        ORAConn.ConnectionString = strConnection
        Try
            ORAConn.Open()
            Debug.Print("사용자 DB Connection")
            Return ORAConn
        Catch ex As Exception
            Debug.Print(ex.Message)
            Console.WriteLine("ERROR - ORCALCE DB CONNCECT : " & ex.Message.ToString & " <<<<<")
            Return ORAConn
        End Try

    End Function

    Public Function _IPPlusConnect() As SqlConnection

        Dim SQLConn As New SqlConnection
        Dim strCategory As String = "DB(IPPlus) Connection"
        SQLConn.ConnectionString = "Data Source=192.168.200.8;Initial Catalog=DBN_IPPlus;User ID=sa;Password=ipplus;Connect Timeout=10"

        Try
            SQLConn.Open()
            Debug.Print("IPPlus DB Connection")
            Return SQLConn
        Catch ex As Exception
            Console.WriteLine("SQLConnect ERROR : " & ex.Message)
            oSysOp.SendMail_forError(strCategory, "DB Connection Error " & ex.Message)
            Return SQLConn
            End
        End Try

    End Function

    Public Function GetLic_Info(ByVal DS As DataSet, ByVal LicInfoDT As DataTable) As DataTable

        'On Error Resume Next
        Dim strCategory As String = "GetLic_Info"
        Try
            Dim MSSqlConn As New SqlConnection
            MSSqlConn = _CAD_SQLConnect()
            If MSSqlConn.State = Data.ConnectionState.Open Then
                Dim SQLSelectComm As New SqlCommand
                SQLSelectComm.Connection = MSSqlConn

                Dim SQlApt As New SqlDataAdapter("SELECT * FROM RNDM_CAD_LIC_INFO WHERE USED=1", MSSqlConn)
                SQlApt.Fill(DS, LicInfoDT.TableName)
                SQlApt.Dispose()

                SQLSelectComm.Dispose()
                MSSqlConn.Close()
                MSSqlConn.Dispose()
            Else
                Console.WriteLine("ERROR - GetLic_Info : DB Connect Error <<<<<")
                LicInfoDT = Nothing
                End
            End If
        Catch ex As Exception
            Console.WriteLine("ERROR - GetLic_Info : " & ex.Message.ToString & " <<<<<")
            oSysOp.SendMail_forError(strCategory, "GetLic_Info Error : " & ex.Message)
            End
        End Try

        Return LicInfoDT

    End Function

    Public Function GetLic_SRV(ByVal DS As DataSet, ByVal LicSrvDT As DataTable) As DataTable

        'On Error Resume Next
        Dim MSSqlConn As New SqlConnection
        MSSqlConn = _CAD_SQLConnect()
        Dim SQLSelectComm As New SqlCommand
        SQLSelectComm.Connection = MSSqlConn
        Dim strCategory As String = "GetLic_SRV"
        Try
            Dim SQlApt As New SqlDataAdapter("SELECT * FROM RNDM_CAD_LIC_SRV", MSSqlConn)
            SQlApt.Fill(DS, LicSrvDT.TableName)
            SQlApt.Dispose()
            SQLSelectComm.Dispose()
            MSSqlConn.Close()
            MSSqlConn.Dispose()
            LicSrvDT.Clear()
        Catch ex As Exception
            Console.WriteLine("GetLic_SRV ERROR : " & ex.Message)
            oSysOp.SendMail_forError(strCategory, "GetLic_SRV Error : " & ex.Message)
            End
        End Try
        Return LicSrvDT
    End Function

    Public Sub DBUpdate(ByVal DS As DataSet, ByVal Lic_InfoDT As DataTable, ByVal Lic_SrvDT As DataTable)

        Dim User_InfoDT As DataTable
        Try
            User_InfoDT = DS.Tables.Add("User_InfoDT")
        Catch ex As Exception
            User_InfoDT = DS.Tables.Item("User_InfoDT")
            User_InfoDT.Clear()
            Console.WriteLine("     >>User_InfoDT Clear!!!")
        End Try

        Dim CAD_SqlConn As New SqlConnection
        CAD_SqlConn = _CAD_SQLConnect()

        'Dim SQLComm As New SqlCommand
        'SQLComm.Connection = CAD_SqlConn

        '==========IP Plus 아이피 뷰 테이블 가져오기===========
        Dim IPPlusSqlConn As New SqlConnection
        IPPlusSqlConn = _CAD_SQLConnect()

        Dim IPPlus_DT As DataTable
        Try
            IPPlus_DT = DS.Tables.Add("IPPlus_DT")
        Catch ex As Exception
            IPPlus_DT = DS.Tables.Item("IPPlus_DT")
            IPPlus_DT.Clear()
            Console.WriteLine("     >>IPPlus_DT Clear!!!")
        End Try

        If IPPlusSqlConn.State = Data.ConnectionState.Open Then
            Dim SQLComm0 As New SqlCommand
            Try
                SQLComm0.Connection = IPPlusSqlConn
            Catch ex As Exception
                Console.WriteLine("     >>SQLCommand ERROR!!!")
                Exit Sub
            End Try

            Dim SelSQL As String = "SELECT * FROM IPListView_Policy"
            Debug.Print(SelSQL)
            Try
                Dim SqlApt As New SqlDataAdapter(SelSQL, IPPlusSqlConn)
                SqlApt.Fill(DS, IPPlus_DT.TableName)
                SqlApt.Dispose()
            Catch ex As Exception
                Console.WriteLine("     >>SqlDataAdapter ERROR!!!")
                SQLComm0.Dispose()
                SQLComm0 = Nothing
                IPPlusSqlConn.Dispose()
                Exit Sub
            End Try
            SQLComm0.Dispose()
            SQLComm0 = Nothing

            IPPlusSqlConn.Close()
            IPPlusSqlConn.Dispose()

        End If
        '==========IP Plus 아이피 뷰 테이블 가져오기===========

        If CAD_SqlConn.State = Data.ConnectionState.Open Then
            Console.WriteLine("     >>MoniteringSystem Connect Success!!!")

            Dim SQLComm1 As New SqlCommand
            Try
                SQLComm1.Connection = CAD_SqlConn
            Catch ex As Exception
                SQLComm1.Dispose()
                Console.WriteLine("     >>SQLCommand ERROR!!!")
                Exit Sub
            End Try

            Dim SelSQL As String = "SELECT * FROM RNDM_CAD_USER_INFO"
            Debug.Print(SelSQL)
            Try
                Dim SqlApt As New SqlDataAdapter(SelSQL, CAD_SqlConn)
                SqlApt.Fill(DS, User_InfoDT.TableName)
                SqlApt.Dispose()
            Catch ex As Exception
                Console.WriteLine("     >>SqlDataAdapter ERROR!!!")
                SQLComm1.Dispose()
                CAD_SqlConn.Dispose()
                Exit Sub
            End Try
            SQLComm1.Dispose()
            Console.WriteLine("     >>SQLCommand Success!!!, Get User Information")


            For index As Integer = 0 To Lic_SrvDT.Rows.Count - 1
                Dim Lic_SrvDTRow As DataRow = Lic_SrvDT.Rows(index)
                Dim UserID As String = Lic_SrvDTRow.Item("USER_ID")
                UserConfirm(Lic_SrvDTRow, User_InfoDT, IPPlus_DT, UserID, False)
            Next

            'RNDM_CAD_LIC_SRV Update
            Dim SQLComm3 As New SqlCommand
            SQLComm3.Connection = CAD_SqlConn

            Console.WriteLine("     1)CAD_LIC_SRV Update 시작")
            Dim DelSql As String = "Delete From RNDM_CAD_LIC_SRV"

            'Dim DelSql As String = "Delete From RNDM_CAD_LIC_SRV_TEST"
            SQLComm3.CommandText = DelSql

            Try
                SQLComm3.ExecuteNonQuery()
            Catch ex As Exception
                Console.WriteLine("ERROR - RNDM_CAD_LIC_SRV Delete :" & ex.Message)
                oSysOp.SendMail_forError("RNDM_CAD_LIC_SRV Delete", "RNDM_CAD_LIC_SRV Delete Error : " & ex.Message)
                End
            End Try
            SQLComm3.Dispose()
            Try
                'Dim SQLComm4 As New SqlCommand
                'SQLComm4.Connection = CAD_SqlConn
                Dim sbMyCopy As New SqlBulkCopy(CAD_SqlConn)
                sbMyCopy.DestinationTableName = "dbo.RNDM_CAD_LIC_SRV"
                sbMyCopy.BulkCopyTimeout = 60000
                sbMyCopy.BatchSize = Lic_SrvDT.Rows.Count
                sbMyCopy.WriteToServer(Lic_SrvDT)
                sbMyCopy.Close()
                'For index As Integer = 0 To Lic_SrvDT.Rows.Count - 1
                '    Dim InsertCmd As String = "INSERT INTO RNDM_CAD_LIC_SRV(CAD_TYPE, LIC_TYPE, LIC_VER, LIC_SRV, USER_ID, NAME, DEPT_ID, DEPT_NAME, IP_INFO, BIZ_CD, M_ORG_NM, C_ORG_CD, C_ORG_NM, LIC_NATION, LIC_DATE, LIC_PROPERTY, TOKEN_USED)VALUES  " & _
                '                                "('" & _
                '                                Lic_SrvDT.Rows(index).Item("CAD_TYPE") & "','" & _
                '                                Lic_SrvDT.Rows(index).Item("LIC_TYPE") & "','" & _
                '                                Lic_SrvDT.Rows(index).Item("LIC_VER") & "','" & _
                '                                Lic_SrvDT.Rows(index).Item("LIC_SRV") & "','" & _
                '                                Lic_SrvDT.Rows(index).Item("USER_ID") & "','" & _
                '                                Lic_SrvDT.Rows(index).Item("NAME") & "','" & _
                '                                Lic_SrvDT.Rows(index).Item("DEPT_ID") & "','" & _
                '                                Lic_SrvDT.Rows(index).Item("DEPT_NAME") & "','" & _
                '                                Lic_SrvDT.Rows(index).Item("IP_INFO") & "','" & _
                '                                Lic_SrvDT.Rows(index).Item("BIZ_CD") & "','" & _
                '                                Lic_SrvDT.Rows(index).Item("M_ORG_NM") & "','" & _
                '                                Lic_SrvDT.Rows(index).Item("C_ORG_CD") & "','" & _
                '                                Lic_SrvDT.Rows(index).Item("C_ORG_NM") & "','" & _
                '                                Lic_SrvDT.Rows(index).Item("LIC_NATION") & "','" & _
                '                               SortDateTime(Lic_SrvDT.Rows(index).Item("LIC_DATE")) & "','" & _
                '                                Lic_SrvDT.Rows(index).Item("LIC_PROPERTY") & "','" & _
                '                                Lic_SrvDT.Rows(index).Item("TOKEN_USED") & _
                '                                "')"


                '    'Dim InsertCmd As String = "INSERT INTO RNDM_CAD_LIC_SRV_TEST(CAD_TYPE, LIC_TYPE, LIC_VER, LIC_SRV, USER_ID, NAME, DEPT_ID, DEPT_NAME, IP_INFO, BIZ_CD, M_ORG_NM, C_ORG_CD, C_ORG_NM, LIC_NATION, LIC_DATE, LIC_PROPERTY)VALUES  " & _
                '    '                            "('" & _
                '    '                            Lic_SrvDT.Rows(index).Item("CAD_TYPE") & "','" & _
                '    '                            Lic_SrvDT.Rows(index).Item("LIC_TYPE") & "','" & _
                '    '                            Lic_SrvDT.Rows(index).Item("LIC_VER") & "','" & _
                '    '                            Lic_SrvDT.Rows(index).Item("LIC_SRV") & "','" & _
                '    '                            Lic_SrvDT.Rows(index).Item("USER_ID") & "','" & _
                '    '                            Lic_SrvDT.Rows(index).Item("NAME") & "','" & _
                '    '                            Lic_SrvDT.Rows(index).Item("DEPT_ID") & "','" & _
                '    '                            Lic_SrvDT.Rows(index).Item("DEPT_NAME") & "','" & _
                '    '                            Lic_SrvDT.Rows(index).Item("IP_INFO") & "','" & _
                '    '                            Lic_SrvDT.Rows(index).Item("BIZ_CD") & "','" & _
                '    '                            Lic_SrvDT.Rows(index).Item("M_ORG_NM") & "','" & _
                '    '                            Lic_SrvDT.Rows(index).Item("C_ORG_CD") & "','" & _
                '    '                            Lic_SrvDT.Rows(index).Item("C_ORG_NM") & "','" & _
                '    '                            Lic_SrvDT.Rows(index).Item("LIC_NATION") & "','" & _
                '    '                           SortDateTime(Lic_SrvDT.Rows(index).Item("LIC_DATE")) & "','" & _
                '    '                            Lic_SrvDT.Rows(index).Item("LIC_PROPERTY") & _
                '    '                            "')"
                '    Debug.Print(InsertCmd)
                '    SQLComm4.CommandText = InsertCmd

                '    Try
                '        SQLComm4.ExecuteNonQuery()
                '    Catch ex As Exception
                '        Console.WriteLine("ERROR - RNDM_CAD_LIC_SRV Insert :" & ex.Message)
                '        oSysOp.SendMail_TEST("RNDM_CAD_LIC_SRV Insert", "RNDM_CAD_LIC_SRV Insert Error")
                '        End
                '    End Try
                '    SQLComm4.Dispose()
                'Next
            Catch ex As Exception
                Console.WriteLine("ERROR - RNDM_CAD_LIC_SRV Insert :" & ex.Message)
                oSysOp.SendMail_forError("RNDM_CAD_LIC_SRV Insert", "RNDM_CAD_LIC_SRV Insert Error : " & ex.Message)
                End
            End Try
            Console.WriteLine("     1)CAD_LIC_SRV Update 완료")

            'RNDM_CAD_LIC_INFO Update
            Console.WriteLine("     2)CAD_LIC_INFO Update 시작")
            For index As Integer = 0 To Lic_InfoDT.Rows.Count - 1
                Dim SQLComm5 As New SqlCommand
                SQLComm5.Connection = CAD_SqlConn

                Dim iNum As Integer
                Dim iOnline As Integer
                Dim iOffline As Integer

                Dim GetRows As DataRow
                GetRows = Lic_InfoDT.Rows(index)

                Try
                    iNum = GetRows.Item("LIC_NUM")
                Catch ex As Exception
                    iNum = "99999"
                End Try
                Try
                    iOnline = GetRows.Item("LIC_ONLINE")
                Catch ex As Exception
                    iOnline = "99999"
                End Try
                Try
                    iOffline = GetRows.Item("LIC_OFFLINE")
                Catch ex As Exception
                    iOffline = "99999"
                End Try

                Debug.Print(GetRows.Item(1) & "/" & GetRows.Item(2) & " : " & iNum & "    " & iOnline & "     " & iOffline)

                Dim CurrentTime As String = SortDateTime(My.Computer.Clock.LocalTime)
                Dim UpdateSql As String = "UPDATE RNDM_CAD_LIC_INFO SET LIC_ONLINE='" & iOnline & "', LIC_OFFLINE='" & iOffline & "', LIC_NUM='" & iNum & "', LIC_DATE=" & "'" & CurrentTime & "'" & _
                " WHERE LIC_TYPE='" & GetRows.Item("LIC_TYPE") & "' AND LIC_VER='" & GetRows.Item("LIC_VER") & "' AND LIC_SRV='" & GetRows.Item("LIC_SRV") & "' AND USED=1"


                Debug.Print(UpdateSql)

                SQLComm5.CommandText = UpdateSql

                Try
                    SQLComm5.ExecuteNonQuery()
                Catch ex As Exception
                    Debug.Print(ex.Message)
                    Console.WriteLine("ERROR - CAD_LIC_INFO UPDATE : " & ex.Message.ToString)
                    oSysOp.SendMail_forError("CAD_LIC_INFO UPDATE", "ERROR - CAD_LIC_INFO UPDATE" & ex.Message.ToString)
                    End
                End Try
                SQLComm5.Dispose()
            Next

            Console.WriteLine("     2)CAD_LIC_INFO Update 완료")
            CAD_SqlConn.Close()
            CAD_SqlConn.Dispose()
        Else
            Console.WriteLine(">>>>>>>>>> DB Insert ERROR!!!")
        End If

    End Sub

    Private Sub UserConfirm(Lic_SrvDTRow As DataRow, User_InfoDT As DataTable, IPPlus_DT As DataTable, UserID As String, Optional IPPlusConfirm As Boolean = False)
        'Dim Lic_SrvDTRow As DataRow = Lic_SrvDT.Rows(index)
        Dim User_InfoDTRow As DataRow()
        Try
            User_InfoDTRow = User_InfoDT.Select("USER_ID='" & UserID & "'")
            If User_InfoDTRow.Count = 1 Then
                Lic_SrvDTRow.Item("NAME") = User_InfoDTRow(0).Item("USER_NAME")
                Lic_SrvDTRow.Item("DEPT_ID") = User_InfoDTRow(0).Item("DEPT_ID")
                Lic_SrvDTRow.Item("DEPT_NAME") = User_InfoDTRow(0).Item("DEPT_NAME")
                Lic_SrvDTRow.Item("BIZ_CD") = User_InfoDTRow(0).Item("BIZ_CD")
                Lic_SrvDTRow.Item("M_ORG_NM") = User_InfoDTRow(0).Item("M_ORG_NM")
                Lic_SrvDTRow.Item("C_ORG_CD") = User_InfoDTRow(0).Item("C_ORG_CD")
                Lic_SrvDTRow.Item("C_ORG_NM") = User_InfoDTRow(0).Item("C_ORG_NM")
            Else

                SetUnknownUser(Lic_SrvDTRow, "-", "-", "-", "-", "-", "-", "-")

                Dim arrUserInfo As New ArrayList
                Dim UnKnownUser_IP As String = Lic_SrvDTRow.Item("IP_INFO")
                Dim UnknownUser_DTRow As DataRow() = Nothing
                If UnKnownUser_IP <> "-" Then
                    UnknownUser_DTRow = User_InfoDT.Select("IP_INFO LIKE '%" & UnKnownUser_IP & "%'")
                    'Else
                    '    GoTo UnKnownUser
                End If

                If UnknownUser_DTRow.Count = 0 Then
                    If IPPlusConfirm = True Then
                        Dim strCategory As String = "사용자 등록"
                        Dim strContents As String = Nothing
                        strContents = "********** IP Plus 확인 User **********" & vbNewLine
                        strContents = strContents & UserID & vbNewLine
                        strContents = strContents & Lic_SrvDTRow.Item("NAME") & vbNewLine
                        strContents = strContents & Lic_SrvDTRow.Item("DEPT_NAME") & vbNewLine
                        strContents = strContents & Lic_SrvDTRow.Item("IP_INFO") & vbNewLine
                        strContents = strContents & Lic_SrvDTRow.Item("CAD_TYPE") & vbNewLine
                        strContents = strContents & Lic_SrvDTRow.Item("LIC_TYPE") & vbNewLine
                        strContents = strContents & Lic_SrvDTRow.Item("LIC_VER") & vbNewLine
                        strContents = strContents & Lic_SrvDTRow.Item("LIC_SRV") & vbNewLine
                        strContents = strContents & "********** 사용자 등록하시기 바랍니다 **********"
                        oSysOp.SendMail(strCategory, strContents)
                        Exit Sub
                    Else
                        '======================== 알수 없는 사용자 분류 =================================================
                        '1. 사용자 분류
                        If Lic_SrvDTRow.Item("USER_ID") = "OFFLINE_USER" Then
                            Dim strLIC_SRV As String = Lic_SrvDTRow.Item("LIC_SRV")
                            If strLIC_SRV = "203.250.10.24" Or strLIC_SRV = "203.250.10.25" Then    '>>금형기술팀
                                SetUnknownUser(Lic_SrvDTRow, "-", "M140000", "금형기술팀", "P010", "진량", "710000", "생산기술센타")
                            ElseIf strLIC_SRV = "210.105.188.204" Then                               '>>중국엔지니어링센타
                                SetUnknownUser(Lic_SrvDTRow, "-", "V1123311", "중국엔지니어링센타", "-", "-", "-", "-")
                            Else
                                SetUnknownUser(Lic_SrvDTRow, "-", "A200001", "연구개발본부", "-", "-", "-", "-")
                            End If
                        ElseIf Lic_SrvDTRow.Item("USER_ID") = "china_ss" Or Lic_SrvDTRow.Item("USER_ID") = "China_ss" Then
                            SetUnknownUser(Lic_SrvDTRow, "-", "C0000010112", "상해삼립회중", "-", "-", "-", "-")
                        ElseIf Lic_SrvDTRow.Item("USER_ID") = "china_dp" Or Lic_SrvDTRow.Item("USER_ID") = "China_dp" Then
                            SetUnknownUser(Lic_SrvDTRow, "-", "C0000010231", "십언동풍삼립차등", "-", "-", "-", "-")
                        ElseIf Lic_SrvDTRow.Item("USER_ID") = "china_yd" Then
                            SetUnknownUser(Lic_SrvDTRow, "-", "C0000010114", "SL연대", "-", "-", "-", "-")
                        ElseIf Lic_SrvDTRow.Item("USER_ID") = "china_bk" Then
                            SetUnknownUser(Lic_SrvDTRow, "-", "C0000010113", "SL북경삼립차등", "-", "-", "-", "-")
                        ElseIf Lic_SrvDTRow.Item("USER_ID") = "china_sl" Then
                            SetUnknownUser(Lic_SrvDTRow, "-", "V1123311", "중국엔지니어링센타", "-", "-", "-", "-")
                        Else
                            'Dim IPPlusIP_DTRow As DataRow()
                            'IPPlusIP_DTRow = IPPlus_DT.Select("IP_String = '" & UnKnownUser_IP & "' AND State = 1")
                            'If IPPlusIP_DTRow.Count = 1 Then
                            '    Dim IPPlus_ID As String = IPPlusIP_DTRow(0).Item("UserField1")
                            '    If IPPlus_ID <> "" Then
                            '        UserConfirm(Lic_SrvDTRow, User_InfoDT, IPPlus_DT, IPPlus_ID, True)
                            '    End If
                            'End If
                        End If

                        '2  Alias, ALM 할당
                        If Lic_SrvDTRow.Item("CAD_TYPE") = "ALIAS" Then
                            SetUnknownUser(Lic_SrvDTRow, "-", "A615000", "디자인팀", "-", "-", "-", "-")
                        ElseIf Lic_SrvDTRow.Item("CAD_TYPE") = "ALM" Then
                            SetUnknownUser(Lic_SrvDTRow, "-", "A210000", "전자개발센타", "-", "-", "-", "-")
                        End If
                        '======================================================================================================
                    End If
                ElseIf UnknownUser_DTRow.Count = 1 Then
                    Lic_SrvDTRow.Item("USER_ID") = UnknownUser_DTRow(0).Item("USER_ID")
                    Lic_SrvDTRow.Item("NAME") = UnknownUser_DTRow(0).Item("USER_NAME")
                    Lic_SrvDTRow.Item("DEPT_ID") = UnknownUser_DTRow(0).Item("DEPT_ID")
                    Lic_SrvDTRow.Item("DEPT_NAME") = UnknownUser_DTRow(0).Item("DEPT_NAME")
                    Lic_SrvDTRow.Item("BIZ_CD") = UnknownUser_DTRow(0).Item("BIZ_CD")
                    Lic_SrvDTRow.Item("M_ORG_NM") = UnknownUser_DTRow(0).Item("M_ORG_NM")
                    Lic_SrvDTRow.Item("C_ORG_CD") = UnknownUser_DTRow(0).Item("C_ORG_CD")
                    Lic_SrvDTRow.Item("C_ORG_NM") = UnknownUser_DTRow(0).Item("C_ORG_NM")
                Else
                    'Dim bMultiIP As Boolean = False
                    'For index1 As Integer = 0 To UnknownUser_DTRow.Count - 1
                    '    Dim arrIP As New ArrayList(Split(UnknownUser_DTRow(index1).Item("IP_INFO"), ":"))
                    '    If arrIP.Contains(UnKnownUser_IP) = True Then
                    '        Lic_SrvDTRow.Item("USER_ID") = UnknownUser_DTRow(index1).Item("USER_ID")
                    '        Lic_SrvDTRow.Item("NAME") = UnknownUser_DTRow(index1).Item("USER_NAME")
                    '        Lic_SrvDTRow.Item("DEPT_ID") = UnknownUser_DTRow(index1).Item("DEPT_ID")
                    '        Lic_SrvDTRow.Item("DEPT_NAME") = UnknownUser_DTRow(index1).Item("DEPT_NAME")
                    '        Lic_SrvDTRow.Item("BIZ_CD") = UnknownUser_DTRow(index1).Item("BIZ_CD")
                    '        Lic_SrvDTRow.Item("M_ORG_NM") = UnknownUser_DTRow(index1).Item("M_ORG_NM")
                    '        Lic_SrvDTRow.Item("C_ORG_CD") = UnknownUser_DTRow(index1).Item("C_ORG_CD")
                    '        Lic_SrvDTRow.Item("C_ORG_NM") = UnknownUser_DTRow(index1).Item("C_ORG_NM")
                    '        bMultiIP = True
                    '        Exit For
                    '    End If
                    'Next
                    'If bMultiIP = False Then

                    '    SetUnknownUser(Lic_SrvDTRow, "-", "-", "-", "-", "-", "-", "-")

                    '    Dim IPPlusIP_DTRow As DataRow()
                    '    IPPlusIP_DTRow = IPPlus_DT.Select("IP_String = '" & UnKnownUser_IP & "' AND State = 1")
                    '    If IPPlusIP_DTRow.Count = 1 Then
                    '        Dim IPPlus_ID As String = IPPlusIP_DTRow(0).Item("UserField1")
                    '        If IPPlus_ID <> "" Then
                    '            UserConfirm(Lic_SrvDTRow, User_InfoDT, IPPlus_DT, IPPlus_ID)
                    '        End If
                    '    End If
                    'End If
                End If
            End If
        Catch ex As Exception
            Console.WriteLine(ex.Message)
        End Try

    End Sub

    'Private Sub SetCADLicDT(ByVal CADLicDTRow As DataRow, ByVal UserInfoValue As ArrayList)

    '    CADLicDTRow.Item("NAME") = UserInfoValue.Item(0)
    '    CADLicDTRow.Item("DEPT_ID") = UserInfoValue.Item(1)
    '    CADLicDTRow.Item("DEPT_NAME") = UserInfoValue.Item(2)
    '    CADLicDTRow.Item("BIZ_CD") = UserInfoValue.Item(3)
    '    CADLicDTRow.Item("M_ORG_NM") = UserInfoValue.Item(4)
    '    CADLicDTRow.Item("C_ORG_CD") = UserInfoValue.Item(5)
    '    CADLicDTRow.Item("C_ORG_NM") = UserInfoValue.Item(6)

    'End Sub

    Public Sub InsertData(ByVal BatchTime As Date)

        Dim DS As New DataSet
        Dim CurrentTime As String = SortDateTime(BatchTime)
        Dim CAD_SqlConn As New SqlConnection
        CAD_SqlConn = _CAD_SQLConnect()

        If CAD_SqlConn.State = ConnectionState.Open Then
            Console.WriteLine("라이선스 정보 입력 상태 : SQLConnection State(" & CAD_SqlConn.State & ")")
            Dim SQLComm As New SqlCommand
            SQLComm.Connection = CAD_SqlConn

            'Dim SqlCmd As String = ""
            Dim SQlApt As New SqlDataAdapter("Select CAD_TYPE, LIC_TYPE, LIC_VER, USER_ID, NAME, DEPT_NAME, IP_INFO, BIZ_CD, M_ORG_NM, C_ORG_CD, C_ORG_NM, LIC_NATION, LIC_PROPERTY, DEPT_ID, LIC_SRV, TOKEN_USED from RNDM_CAD_LIC_SRV", CAD_SqlConn)
            Dim LicSrvDT As DataTable
            Dim LicUseDT As DataTable
            Dim LicUseNumforMax As DataTable

            Try
                LicSrvDT = DS.Tables.Add("LICSrcInfo")
            Catch ex As Exception
                LicSrvDT = DS.Tables.Item("LICSrcInfo")
                LicSrvDT.Clear()
            End Try

            Try
                SQlApt.Fill(DS, "LICSrcInfo")
                SQlApt.Dispose()
            Catch ex As Exception
                oSysOp.SendMail_forError("LICSrcInfoForRNDM_CAD_LIC_USE", "ERROR-LICSrcInfoForRNDM_CAD_LIC_USE" & ex.Message)
                End
            End Try

            Dim SQLAPTforNumMax As New SqlDataAdapter("SELECT MAX(NUM) FROM RNDM_CAD_LIC_USE", CAD_SqlConn)

            Try
                LicUseNumforMax = DS.Tables.Add("LICNumMax")
            Catch ex As Exception
                LicUseNumforMax = DS.Tables.Item("LICNumMax")
                LicUseNumforMax.Clear()
            End Try

            Try
                SQLAPTforNumMax.Fill(DS, "LICNumMax")
                SQLAPTforNumMax.Dispose()
            Catch ex As Exception
                End
            End Try

            Dim intLicUseNumMax As Integer
            intLicUseNumMax = CType(LicUseNumforMax.Rows.Item(0).Item(0), Integer)

            LicUseDT = DS.Tables.Add("LICUseInfo")
            LicUseDT.Columns.Add("NUM")
            LicUseDT.Columns.Add("CAD_TYPE")
            LicUseDT.Columns.Add("LIC_TYPE")
            LicUseDT.Columns.Add("LIC_VER")
            LicUseDT.Columns.Add("USER_ID")
            LicUseDT.Columns.Add("NAME")
            LicUseDT.Columns.Add("DEPT_NAME")
            LicUseDT.Columns.Add("IP_INFO")
            LicUseDT.Columns.Add("COMP_NAME")
            LicUseDT.Columns.Add("USE_DATE")
            LicUseDT.Columns.Add("BIZ_CD")
            LicUseDT.Columns.Add("M_ORG_NM")
            LicUseDT.Columns.Add("C_ORG_CD")
            LicUseDT.Columns.Add("C_ORG_NM")
            LicUseDT.Columns.Add("LIC_PROPERTY")
            LicUseDT.Columns.Add("LIC_NATION")
            LicUseDT.Columns.Add("LIC_SRV")
            LicUseDT.Columns.Add("DEPT_ID")
            LicUseDT.Columns.Add("TOKEN_USED")

            For index As Integer = 0 To LicSrvDT.Rows.Count - 1
                Dim NewRow As DataRow
                intLicUseNumMax = intLicUseNumMax + 1
                NewRow = LicUseDT.Rows.Add
                NewRow.Item("NUM") = intLicUseNumMax
                NewRow.Item("CAD_TYPE") = LicSrvDT.Rows(index).Item("CAD_TYPE")
                NewRow.Item("LIC_TYPE") = LicSrvDT.Rows(index).Item("LIC_TYPE")
                NewRow.Item("LIC_VER") = LicSrvDT.Rows(index).Item("LIC_VER")
                NewRow.Item("USER_ID") = LicSrvDT.Rows(index).Item("USER_ID")
                NewRow.Item("NAME") = LicSrvDT.Rows(index).Item("NAME")
                NewRow.Item("DEPT_NAME") = LicSrvDT.Rows(index).Item("DEPT_NAME")
                NewRow.Item("IP_INFO") = LicSrvDT.Rows(index).Item("IP_INFO")
                NewRow.Item("COMP_NAME") = "-"
                NewRow.Item("USE_DATE") = CurrentTime
                NewRow.Item("BIZ_CD") = LicSrvDT.Rows(index).Item("BIZ_CD")
                NewRow.Item("M_ORG_NM") = LicSrvDT.Rows(index).Item("M_ORG_NM")
                NewRow.Item("C_ORG_CD") = LicSrvDT.Rows(index).Item("C_ORG_CD")
                NewRow.Item("C_ORG_NM") = LicSrvDT.Rows(index).Item("C_ORG_NM")
                NewRow.Item("LIC_PROPERTY") = LicSrvDT.Rows(index).Item("LIC_PROPERTY")
                NewRow.Item("LIC_NATION") = LicSrvDT.Rows(index).Item("LIC_NATION")
                NewRow.Item("LIC_SRV") = LicSrvDT.Rows(index).Item("LIC_SRV")
                NewRow.Item("DEPT_ID") = LicSrvDT.Rows(index).Item("DEPT_ID")
                NewRow.Item("TOKEN_USED") = LicSrvDT.Rows(index).Item("TOKEN_USED")
            Next

            Try
                Dim sbMyCopy As New SqlBulkCopy(CAD_SqlConn)
                sbMyCopy.DestinationTableName = "dbo.RNDM_CAD_LIC_USE"
                sbMyCopy.BulkCopyTimeout = 60000
                sbMyCopy.BatchSize = LicUseDT.Rows.Count
                sbMyCopy.WriteToServer(LicUseDT)
                sbMyCopy.Close()
            Catch ex As Exception
                Debug.Print(ex.Message)
                Console.WriteLine("ERROR : INSERT CAD_LIC_USE : " & ex.Message.ToString)
                oSysOp.SendMail_forError("INSERT CAD_LIC_USE", "ERROR : INSERT CAD_LIC_USE : " & ex.Message.ToString)
                End
            End Try

            'For index As Integer = 0 To LicSrvDT.Rows.Count - 1
            '    Dim InsertCmd As String = "INSERT INTO RNDM_CAD_LIC_USE(CAD_TYPE, LIC_TYPE, LIC_VER, USER_ID, NAME, DEPT_NAME, IP_INFO, USE_DATE, BIZ_CD, M_ORG_NM, C_ORG_CD, C_ORG_NM, LIC_NATION, LIC_PROPERTY, DEPT_ID, LIC_SRV, TOKEN_USED)VALUES  " & _
            '                                "('" & _
            '                                LicSrvDT.Rows(index).Item("CAD_TYPE") & "','" & _
            '                                LicSrvDT.Rows(index).Item("LIC_TYPE") & "','" & _
            '                                LicSrvDT.Rows(index).Item("LIC_VER") & "','" & _
            '                                LicSrvDT.Rows(index).Item("USER_ID") & "','" & _
            '                                LicSrvDT.Rows(index).Item("NAME") & "','" & _
            '                                LicSrvDT.Rows(index).Item("DEPT_NAME") & "','" & _
            '                                LicSrvDT.Rows(index).Item("IP_INFO") & "','" & _
            '                                CurrentTime & "','" & _
            '                                LicSrvDT.Rows(index).Item("BIZ_CD") & "','" & _
            '                                LicSrvDT.Rows(index).Item("M_ORG_NM") & "','" & _
            '                                LicSrvDT.Rows(index).Item("C_ORG_CD") & "','" & _
            '                                LicSrvDT.Rows(index).Item("C_ORG_NM") & "','" & _
            '                                LicSrvDT.Rows(index).Item("LIC_NATION") & "','" & _
            '                                LicSrvDT.Rows(index).Item("LIC_PROPERTY") & "','" & _
            '                                LicSrvDT.Rows(index).Item("DEPT_ID") & "','" & _
            '                                LicSrvDT.Rows(index).Item("LIC_SRV") & "','" & _
            '                                LicSrvDT.Rows(index).Item("TOKEN_USED") & _
            '                                "')"
            '    Debug.Print(InsertCmd)
            '    Try
            '        SQLComm.CommandText = InsertCmd
            '        SQLComm.CommandTimeout = 10000
            '        SQLComm.ExecuteNonQuery()
            '    Catch ex As Exception
            '        Debug.Print(ex.Message)
            '        Console.WriteLine("ERROR : INSERT CAD_LIC_USE : " & ex.Message.ToString)
            '        oSysOp.SendMail_TEST("INSERT CAD_LIC_USE", "ERROR : INSERT CAD_LIC_USE : " & ex.Message.ToString)
            '        End
            '    End Try
            'Next


            '매일 12시에 라이선스 정보 기록 및 EP 정보 싱크
            If BatchTime.Hour = 12 Then
                Console.WriteLine("======= Insert License Information Start =======")
                InsertLicInfo(DS, CAD_SqlConn, SQLComm, CurrentTime)
                Console.WriteLine("======= Insert License Information End   =======")

                Console.WriteLine("======= CAD Portal UserInfo Sync. Start =======")
                UserInformation_Sync(DS, CAD_SqlConn, SQLComm)
                Console.WriteLine("======= CAD Portal UserInfo Sync. End =======")

                Console.WriteLine("======= EP Portal UserInfo Sync. Start =======")
                Interface_UserInfo(CAD_SqlConn, SQLComm, CurrentTime)
                Console.WriteLine("======= EP Portal UserInfo Sync. End =======")

                Console.WriteLine("======= EP Portal Orgranization Sync. Start =======")
                Interface_Organization(CAD_SqlConn, SQLComm, CurrentTime)
                Console.WriteLine("======= EP Portal Orgranization Sync. End =======")
            End If
            SQLComm.Dispose()
        End If

        CAD_SqlConn.Close()
        CAD_SqlConn.Dispose()

    End Sub

    Private Sub InsertLicInfo(ByVal DS As DataSet, ByVal CADSqlConn As SqlConnection, ByVal SQLComm As SqlCommand, ByVal CurrentTime As String)

        Dim SqlCmd As String = ""
        Dim SQlApt As New SqlDataAdapter("SELECT CAD_TYPE, LIC_TYPE, LIC_VER, LIC_SRV, LIC_NUM, LIC_ONLINE, LIC_OFFLINE, LIC_NATION, LIC_PROPERTY, LIC_DESCRIPTION FROM RNDM_CAD_LIC_INFO WHERE USED=1", CADSqlConn)
        Dim LicSrvDT As DataTable
        Try
            LicSrvDT = DS.Tables.Add("LICINFO_DT")
        Catch ex As Exception
            LicSrvDT = DS.Tables.Item("LICINFO_DT")
            LicSrvDT.Clear()
        End Try

        SQlApt.Fill(DS, "LICINFO_DT")
        SQlApt.Dispose()

        For index As Integer = 0 To LicSrvDT.Rows.Count - 1
            Dim InsertCmd As String = "INSERT INTO RNDM_CAD_LIC_INFO_STATES(CAD_TYPE, LIC_TYPE, LIC_VER, LIC_SRV, LIC_NUM, LIC_ONLINE, LIC_OFFLINE, LIC_NATION, LIC_PROPERTY, LIC_DESCRIPTION, CRE_DATE)VALUES  " & _
                                        "('" & _
                                        LicSrvDT.Rows(index).Item("CAD_TYPE") & "','" & _
                                        LicSrvDT.Rows(index).Item("LIC_TYPE") & "','" & _
                                        LicSrvDT.Rows(index).Item("LIC_VER") & "','" & _
                                        LicSrvDT.Rows(index).Item("LIC_SRV") & "','" & _
                                        LicSrvDT.Rows(index).Item("LIC_NUM") & "','" & _
                                        LicSrvDT.Rows(index).Item("LIC_ONLINE") & "','" & _
                                        LicSrvDT.Rows(index).Item("LIC_OFFLINE") & "','" & _
                                        LicSrvDT.Rows(index).Item("LIC_NATION") & "','" & _
                                        LicSrvDT.Rows(index).Item("LIC_PROPERTY") & "','" & _
                                        LicSrvDT.Rows(index).Item("LIC_DESCRIPTION") & "','" & _
                                        CurrentTime & _
                                        "')"


            Debug.Print(InsertCmd)
            SQLComm.CommandText = InsertCmd

            Try
                SQLComm.ExecuteNonQuery()
            Catch ex As Exception
                Debug.Print(ex.Message)
                Console.WriteLine("ERROR - INSERT RNDM_CAD_INFO_STATES : " & ex.Message.ToString)
                Exit Sub
            End Try
        Next

    End Sub

    'EP 정보와 사용자 동기화
    Public Sub UserInformation_Sync(ByVal DS As DataSet, ByVal CADSqlConn As SqlConnection, ByVal SQLComm As SqlCommand)

        Dim UserSyncDT As DataTable
        Try
            UserSyncDT = DS.Tables.Add("UserSync_DT")
        Catch ex As Exception
            UserSyncDT = DS.Tables.Item("UserSync_DT")
            UserSyncDT.Clear()
        End Try

        Dim SQlApt As New SqlDataAdapter("SELECT USER_ID, USER_NAME, DEPT_NAME FROM RNDM_CAD_USER_INFO WHERE CATEGORY='EP'", CADSqlConn)

        SQlApt.Fill(DS, "UserSync_DT")
        SQlApt.Dispose()

        'ep 접속 정보 받아오기
        Dim EPOraConn As New OracleConnection
        EPOraConn = _EP_Connect()
        Dim ORACmd As New OracleCommand

        For index As Integer = 0 To UserSyncDT.Rows.Count - 1

            Dim UserID As String = UserSyncDT.Rows(index).Item(0)
            Dim UserName As String = UserSyncDT.Rows(index).Item(1)
            Dim DeptName As String = UserSyncDT.Rows(index).Item(2)

            Dim strCmd As String
            strCmd = "SELECT USER_NAME, DEPT_ID, DEPT_NAME, BIZ_CD, M_ORG_NM, C_ORG_CD, C_ORG_NM FROM INTERFACE.V_USERINFOLIST WHERE USER_ID='" & UserID & "'"

            ORACmd.Connection = EPOraConn
            ORACmd.CommandText = strCmd

            Dim EPInfo As New ArrayList

            Dim ORAReader As OracleDataReader
            ORAReader = ORACmd.ExecuteReader
            If (ORAReader.HasRows) Then
                Do
                    ORAReader.Read()
                    EPInfo.Add(ORAReader.GetValue(0))
                    EPInfo.Add(ORAReader.GetValue(1))
                    EPInfo.Add(ORAReader.GetValue(2))
                    EPInfo.Add(ORAReader.GetValue(3))
                    EPInfo.Add(ORAReader.GetValue(4))
                    EPInfo.Add(ORAReader.GetValue(5))
                    EPInfo.Add(ORAReader.GetValue(6))
                Loop While ORAReader.IsClosed = True
                ORAReader.Close()
            Else
                Dim DelCmd As String
                DelCmd = "DELETE FROM RNDM_CAD_USER_INFO WHERE USER_ID='" & UserID & "'"

                Debug.Print(DelCmd)
                SQLComm.Connection = CADSqlConn
                SQLComm.CommandText = DelCmd

                Try
                    SQLComm.ExecuteNonQuery()

                    Dim strContents As String = UserID & "/" & UserName & "/" & DeptName & " : 삭제"
                    Console.WriteLine(">>>>>>>>>> 사용자 삭제 : " & strContents)
                    Dim strCategory As String = "사용자 삭제 알림"
                    oSysOp.SendMail(strCategory, strContents)

                Catch ex As Exception
                    Debug.Print(ex.Message)
                    Console.WriteLine("ERROR - DELETE USER INFO : " & ex.Message.ToString)
                End Try

                GoTo PASS
            End If

            Dim _USER_NAME As String = "-"
            Dim _DEPT_ID As String = "-"
            Dim _DEPT_NAME As String = "-"
            Dim _BIZ_CD As String = "-"
            Dim _M_ORG_NM As String = "-"
            Dim _C_ORG_CD As String = "-"
            Dim _C_ORG_NM As String = "-"

            Try
                _USER_NAME = EPInfo(0)
            Catch ex As Exception

            End Try

            Try
                _DEPT_ID = EPInfo(1)
            Catch ex As Exception

            End Try

            Try
                _DEPT_NAME = EPInfo(2)
            Catch ex As Exception

            End Try

            Try
                _BIZ_CD = EPInfo(3)
            Catch ex As Exception

            End Try

            Try
                _M_ORG_NM = EPInfo(4)
            Catch ex As Exception

            End Try
            Try
                _C_ORG_CD = EPInfo(5)
            Catch ex As Exception

            End Try

            Try
                _C_ORG_NM = EPInfo(6)
            Catch ex As Exception

            End Try

            Dim strSQL As String
            strSQL = "UPDATE RNDM_CAD_USER_INFO SET " & _
            "USER_NAME='" & _USER_NAME & "'" & ", " & _
            "DEPT_ID='" & _DEPT_ID & "'" & ", " & _
            "DEPT_NAME='" & _DEPT_NAME & "'" & ", " & _
            "BIZ_CD='" & _BIZ_CD & "'" & ", " & _
            "M_ORG_NM='" & _M_ORG_NM & "'" & ", " & _
            "C_ORG_CD='" & _C_ORG_CD & "'" & ", " & _
            "C_ORG_NM='" & _C_ORG_NM & "' " & _
            "WHERE USER_ID='" & UserID & "'"

            Debug.Print(strSQL)
            SQLComm.CommandText = strSQL

            Try
                SQLComm.ExecuteNonQuery()
            Catch ex As Exception
                Debug.Print(ex.Message)
                Console.WriteLine("ERROR - UPDATE USER_INFO : " & ex.Message.ToString)
            End Try
PASS:
        Next
        ORACmd.Dispose()

        EPOraConn.Close()
        EPOraConn.Dispose()
        EPOraConn = Nothing

    End Sub

    Public Sub Oraganization_Sync(ByVal DS As DataSet, ByVal CADSqlConn As SqlConnection, ByVal SQLComm As SqlCommand)

        '조직도 받아오기
        Dim OrgSyncDT As DataTable
        Try
            OrgSyncDT = DS.Tables.Add("OrgSync_DT")
        Catch ex As Exception
            OrgSyncDT = DS.Tables.Item("OrgSync_DT")
            OrgSyncDT.Clear()
        End Try

        Dim SQlApt As New SqlDataAdapter("SELECT * FROM RNDM_CAD_LIC_ORGANIZATIONINFO", CADSqlConn)

        SQlApt.Fill(DS, OrgSyncDT.TableName)
        SQlApt.Dispose()

        '=================ep 조직도 받아 오기(Oracle)================================
        Dim EPOrgSyncDT As DataTable
        Try
            EPOrgSyncDT = DS.Tables.Add("EPOrgSync_DT")
        Catch ex As Exception
            EPOrgSyncDT = DS.Tables.Item("EPOrgSync_DT")
            EPOrgSyncDT.Clear()
        End Try

        Dim EPOraConn As New OracleConnection
        EPOraConn = _EP_Connect()

        Dim ORAApt As New OracleDataAdapter("SELECT * FROM INTERFACE.V_ORGANIZATIONINFO_ALL WHERE(is_deleted = 0) ORDER BY ORG_NAME", EPOraConn)

        ORAApt.Fill(DS, EPOrgSyncDT.TableName)
        ORAApt.Dispose()

        EPOraConn.Close()
        EPOraConn.Dispose()
        EPOraConn = Nothing
        '============================================================================

        'ORG_ID 정리
        'CAD Portal 관리 조직
        Dim arrORG_ID_CAD As New ArrayList
        Dim arrORG_ID_EP As New ArrayList

        For index As Integer = 0 To OrgSyncDT.Rows.Count - 1
            arrORG_ID_CAD.Add(OrgSyncDT.Rows(index).Item(3))
        Next

        For index As Integer = 0 To EPOrgSyncDT.Rows.Count - 1
            arrORG_ID_EP.Add(EPOrgSyncDT.Rows(index).Item(3))
        Next

        '삭제할 조직 확인
        Dim DetState As Boolean = False
        For index As Integer = 0 To arrORG_ID_CAD.Count - 1
            If arrORG_ID_EP.Contains(arrORG_ID_CAD.Item(index)) = False Then
                If OrgSyncDT.Rows(index).Item(11) <> 0 Or OrgSyncDT.Rows(index).Item(12) <> 0 Or OrgSyncDT.Rows(index).Item(13) <> 0 & _
                OrgSyncDT.Rows(index).Item(14) <> 0 Or OrgSyncDT.Rows(index).Item(15) <> 0 Or OrgSyncDT.Rows(index).Item(16) <> 0 Then
                    DetState = True
                    Dim strCagegory As String = "조직 삭제 알림"
                    Dim strContents As String
                    strContents = OrgSyncDT.Columns(0).ColumnName & " : " & OrgSyncDT.Rows(index).Item(0) & vbNewLine & _
                    OrgSyncDT.Columns(3).ColumnName & " : " & OrgSyncDT.Rows(index).Item(3) & vbNewLine & _
                    OrgSyncDT.Columns(11).ColumnName & " : " & OrgSyncDT.Rows(index).Item(11) & vbNewLine & _
                    OrgSyncDT.Columns(12).ColumnName & " : " & OrgSyncDT.Rows(index).Item(12) & vbNewLine & _
                    OrgSyncDT.Columns(13).ColumnName & " : " & OrgSyncDT.Rows(index).Item(13) & vbNewLine & _
                    OrgSyncDT.Columns(14).ColumnName & " : " & OrgSyncDT.Rows(index).Item(14) & vbNewLine & _
                    OrgSyncDT.Columns(15).ColumnName & " : " & OrgSyncDT.Rows(index).Item(15) & vbNewLine & _
                    OrgSyncDT.Columns(16).ColumnName & " : " & OrgSyncDT.Rows(index).Item(16) & vbNewLine & _
                    "라이선스가 할당된 조직이 삭제 되었습니다. 서버 확인 바랍니다" & vbNewLine
                    Console.WriteLine(">>>>>>>>>> 조직 삭제 : " & strCagegory)
                    oSysOp.SendMail(strCagegory, strContents)
                End If
                OrgSyncDT.Rows(index).Delete()
            End If
        Next

        '추가할 조직 확인
        Dim AddState As Boolean = False
        For index As Integer = 0 To arrORG_ID_EP.Count - 1
            If arrORG_ID_CAD.Contains(arrORG_ID_EP.Item(index)) = False Then
                AddState = True
                Dim SetDataRow As DataRow
                SetDataRow = OrgSyncDT.Rows.Add
                SetDataRow("ORG_NAME") = EPOrgSyncDT.Rows(index).Item("ORG_NAME")
                SetDataRow("ORG_OTHER_NAME") = EPOrgSyncDT.Rows(index).Item("ORG_OTHER_NAME")
                SetDataRow("ORG_ABBR_NAME") = EPOrgSyncDT.Rows(index).Item("ORG_ABBR_NAME")
                SetDataRow("ORG_ID") = EPOrgSyncDT.Rows(index).Item("ORG_ID")
                SetDataRow("ORG_CODE") = EPOrgSyncDT.Rows(index).Item("ORG_CODE")
                SetDataRow("ORG_PARENT_ID") = EPOrgSyncDT.Rows(index).Item("ORG_PARENT_ID")
                SetDataRow("ORG_ORDER") = EPOrgSyncDT.Rows(index).Item("ORG_ORDER")
                SetDataRow("SERVERS") = EPOrgSyncDT.Rows(index).Item("SERVERS")
                SetDataRow("ORG_TYPE") = EPOrgSyncDT.Rows(index).Item("ORG_TYPE")
                SetDataRow("DESCRIPTION") = EPOrgSyncDT.Rows(index).Item("DESCRIPTION")
                SetDataRow("COMPANY_ID") = EPOrgSyncDT.Rows(index).Item("COMPANY_ID")
                SetDataRow("CAT_ONLINE_LIC") = 0
                SetDataRow("CAT_OFFLINE_LIC") = 0
                SetDataRow("NX_ONLINE_LIC") = 0
                SetDataRow("NX_OFFLINE_LIC") = 0
                SetDataRow("APR_ONLINE_LIC") = 0
                SetDataRow("APR_OFFLINE_LIC") = 0

                Dim strCagegory As String = "조직 추가 알림"
                Dim strContents As String
                strContents = EPOrgSyncDT.Columns(0).ColumnName & " : " & EPOrgSyncDT.Rows(index).Item(0) & vbNewLine & _
                EPOrgSyncDT.Columns(3).ColumnName & " : " & EPOrgSyncDT.Rows(index).Item(3) & vbNewLine & _
                "조직이 추가 되었습니다. 서버 확인 바랍니다." & vbNewLine
                Debug.Print(strContents)
                oSysOp.SendMail(strCagegory, strContents)
            End If
        Next

        If DetState = True Or AddState = True Then
            If CADSqlConn.State = ConnectionState.Open Then
                Dim DelSql As String = "Delete From RNDM_CAD_LIC_ORGANIZATIONINFO"

                SQLComm.CommandText = DelSql

                Try
                    SQLComm.ExecuteNonQuery()
                Catch ex As Exception
                    Console.WriteLine("ERROR - RNDM_CAD_LIC_SRV Delete :" & ex.Message)
                End Try

                For index As Integer = 0 To OrgSyncDT.Rows.Count - 1
                    Dim InsertCmd As String = "INSERT INTO RNDM_CAD_LIC_ORGANIZATIONINFO(ORG_NAME, ORG_OTHER_NAME, ORG_ABBR_NAME, ORG_ID, ORG_CODE, ORG_PARENT_ID, ORG_ORDER, SERVERS, ORG_TYPE, DESCRIPTION, COMPANY_ID, CAT_ONLINE_LIC, CAT_OFFLINE_LIC, NX_ONLINE_LIC, NX_OFFLINE_LIC, APR_ONLINE_LIC, APR_OFFLINE_LIC)VALUES  " & _
                                "('" & _
                                OrgSyncDT.Rows(index).Item("ORG_NAME") & "','" & _
                                OrgSyncDT.Rows(index).Item("ORG_OTHER_NAME") & "','" & _
                                OrgSyncDT.Rows(index).Item("ORG_ABBR_NAME") & "','" & _
                                OrgSyncDT.Rows(index).Item("ORG_ID") & "','" & _
                                OrgSyncDT.Rows(index).Item("ORG_CODE") & "','" & _
                                OrgSyncDT.Rows(index).Item("ORG_PARENT_ID") & "','" & _
                                OrgSyncDT.Rows(index).Item("ORG_ORDER") & "','" & _
                                OrgSyncDT.Rows(index).Item("SERVERS") & "','" & _
                                OrgSyncDT.Rows(index).Item("ORG_TYPE") & "','" & _
                                OrgSyncDT.Rows(index).Item("DESCRIPTION") & "','" & _
                                OrgSyncDT.Rows(index).Item("COMPANY_ID") & "','" & _
                                OrgSyncDT.Rows(index).Item("CAT_ONLINE_LIC") & "','" & _
                                OrgSyncDT.Rows(index).Item("CAT_OFFLINE_LIC") & "','" & _
                                OrgSyncDT.Rows(index).Item("NX_ONLINE_LIC") & "','" & _
                                OrgSyncDT.Rows(index).Item("NX_OFFLINE_LIC") & "','" & _
                                OrgSyncDT.Rows(index).Item("APR_ONLINE_LIC") & "','" & _
                                OrgSyncDT.Rows(index).Item("APR_OFFLINE_LIC") & _
                                "')"

                    Debug.Print(InsertCmd)
                    SQLComm.CommandText = InsertCmd

                    Try
                        SQLComm.ExecuteNonQuery()
                    Catch ex As Exception
                        Debug.Print(ex.Message)
                        Console.WriteLine("ERROR - INSERT ORGANIZATION : " & ex.Message.ToString)
                    End Try
                Next
            End If
        End If

    End Sub

    Private Function SortDateTime(ByVal CurrentTime As Date) As String

        Dim SQLDataTime As String
        SQLDataTime = CStr(CurrentTime.Year) + "-" + CStr(CurrentTime.Month) + "-" + CStr(CurrentTime.Day)
        SQLDataTime = SQLDataTime + " " + CStr(CurrentTime.Hour) + ":" + CStr(CurrentTime.Minute) + ":" + CStr(CurrentTime.Second)
        Debug.Print(SQLDataTime)

        Return SQLDataTime

    End Function

    Private Sub SetUnknownUser(ByVal DTRow As DataRow, ByVal NAME As String, ByVal DEPT_ID As String, ByVal DEPT_NAME As String, ByVal BIZ_CD As String, ByVal M_ORG_NM As String, ByVal C_ORG_CD As String, ByVal C_ORG_NM As String)
        DTRow.Item("NAME") = NAME
        DTRow.Item("DEPT_ID") = DEPT_ID
        DTRow.Item("DEPT_NAME") = DEPT_NAME
        DTRow.Item("BIZ_CD") = BIZ_CD
        DTRow.Item("M_ORG_NM") = M_ORG_NM
        DTRow.Item("C_ORG_CD") = C_ORG_CD
        DTRow.Item("C_ORG_NM") = C_ORG_NM
    End Sub

#Region "Portal Interface"

    Public Sub Interface_UserInfo(ByVal CADSqlConn As SqlConnection, ByVal SQLComm As SqlCommand, ModifyDate As String)

        'CAD 서버 사용자 정보 가져오기
        Dim I_UserInfoDT As New DataTable
        Dim SQlApt As New SqlDataAdapter("SELECT * FROM INTERFACE_USERINFO", CADSqlConn)

        SQlApt.Fill(I_UserInfoDT)
        SQlApt.Dispose()

        'EP 포탈 사용자 정보 가져오기
        Dim EPOraConn As New OracleConnection
        EPOraConn = _EP_Connect()
        Dim ORACmd As New OracleCommand

        Dim strSel As String = "SELECT USER_ID, " & _
                                        "USER_NAME, " & _
                                        "USER_OTHER_NAME, " & _
                                        "USER_DISPLAY_NAME, " & _
                                        "USER_UID, " & _
                                        "COMP_ID, " & _
                                        "COMP_NAME, " & _
                                        "DEPT_ID, " & _
                                        "DEPT_NAME, " & _
                                        "BIZ_CD, " & _
                                        "M_ORG_NM, " & _
                                        "C_ORG_CD, " & _
                                        "C_ORG_NM, " & _
                                        "ORG_CD, " & _
                                        "ORGF_NM, " & _
                                        "ORG_DISPLAY_NAME, " & _
                                        "OU_GROUP, " & _
                                        "GRADE_CODE, " & _
                                        "TITLE_CODE, " & _
                                        "GRADE_NAME, " & _
                                        "TITLE_NAME, " & _
                                        "USER_ORDER, " & _
                                        "EMPLOYEE_ID, " & _
                                        "SYSMAIL " & _
                                            "FROM INTERFACE.V_USERINFOLIST " & _
                                            "WHERE GRADE_CODE <>200"
        Debug.Print(strSel)

        Dim ORAApt As New OracleDataAdapter(strSel, EPOraConn)
        Dim HRInfo_DT As New DataTable
        ORAApt.Fill(HRInfo_DT)
        ORAApt.Dispose()

        ORACmd.Dispose()

        EPOraConn.Close()
        EPOraConn.Dispose()
        EPOraConn = Nothing

        '퇴사자 삭제
        For Each Row In I_UserInfoDT.AsEnumerable()
            Dim strUserId As String = Row.Field(Of String)("USER_ID")

            'HR 테이블에서 아이디 찾기
            Dim GetID_Query = From IDRow In HRInfo_DT.AsEnumerable() _
                                Where IDRow.Field(Of String)("USER_ID") = strUserId _
                                Select IDRow


            'getid_query <> nothing 일 경우 모든 정보 업데이트
            If GetID_Query.Count <> 0 Then
                UpdateUesrInfo(CADSqlConn, SQLComm, GetID_Query(0))
                'getid_query 가 nothing 일 경우 퇴사자 처리
            Else
                Debug.Print(strUserId)
                '유저 퇴사자 테이블에 추가
                InsertUserInfo(2, CADSqlConn, SQLComm, Row, ModifyDate)

                '유저 삭제
                DeleteUserInfo(CADSqlConn, SQLComm, Row)
            End If
        Next

        '입사자 추가
        For Each Row In HRInfo_DT.AsEnumerable()
            Dim strUserId As String = Row.Field(Of String)("USER_ID")

            'HR 테이블에서 아이디 찾기
            Dim GetID_Query = From IDRow In I_UserInfoDT.AsEnumerable() _
                                Where IDRow.Field(Of String)("USER_ID") = strUserId _
                                Select IDRow

            'getid_query 가 nothing 일 경우 입사자 처리
            If GetID_Query.Count = 0 Then
                '유저 입사자 테이블에 추가
                InsertUserInfo(1, CADSqlConn, SQLComm, Row, ModifyDate)
            End If
        Next

        I_UserInfoDT.Dispose()
        HRInfo_DT.Dispose()

    End Sub

    '1 = INTERFACE_USERINFO
    '2 = RETIRED_USERINFO
    Private Sub InsertUserInfo(i As Integer, ByVal CADSqlConn As SqlConnection, ByVal SQLComm As SqlCommand, Row As DataRow, ModifyDate As String)

        Dim strValues As String = String.Empty
        For index = 0 To Row.ItemArray.Count - 1
            If index = 0 Then
                strValues = "'" & Row.ItemArray(index) & "'"
            Else
                strValues = strValues & ", '" & Row.ItemArray(index) & "'"
            End If
        Next

        Dim InsertSQL As String = String.Empty
        If i = 1 Then
            InsertSQL = "INSERT INTO INTERFACE_USERINFO(" & _
                                        "USER_ID, " & _
                                        "USER_NAME, " & _
                                        "USER_OTHER_NAME, " & _
                                        "USER_DISPLAY_NAME, " & _
                                        "USER_UID, " & _
                                        "COMP_ID, " & _
                                        "COMP_NAME, " & _
                                        "DEPT_ID, " & _
                                        "DEPT_NAME, " & _
                                        "BIZ_CD, " & _
                                        "M_ORG_NM, " & _
                                        "C_ORG_CD, " & _
                                        "C_ORG_NM, " & _
                                        "ORG_CD, " & _
                                        "ORGF_NM, " & _
                                        "ORG_DISPLAY_NAME, " & _
                                        "OU_GROUP, " & _
                                        "GRADE_CODE, " & _
                                        "TITLE_CODE, " & _
                                        "GRADE_NAME, " & _
                                        "TITLE_NAME, " & _
                                        "USER_ORDER, " & _
                                        "EMPLOYEE_ID, " & _
                                        "SYSMAIL, " & _
                                        "MODIFY_DATE" & _
                                        ")" & _
                                        " VALUES(" & strValues & ", '" & ModifyDate & "')"
        Else
            InsertSQL = "INSERT INTO RETIRED_USERINFO(" & _
                                        "USER_ID, " & _
                                        "USER_NAME, " & _
                                        "USER_OTHER_NAME, " & _
                                        "USER_DISPLAY_NAME, " & _
                                        "USER_UID, " & _
                                        "COMP_ID, " & _
                                        "COMP_NAME, " & _
                                        "DEPT_ID, " & _
                                        "DEPT_NAME, " & _
                                        "BIZ_CD, " & _
                                        "M_ORG_NM, " & _
                                        "C_ORG_CD, " & _
                                        "C_ORG_NM, " & _
                                        "ORG_CD, " & _
                                        "ORGF_NM, " & _
                                        "ORG_DISPLAY_NAME, " & _
                                        "OU_GROUP, " & _
                                        "GRADE_CODE, " & _
                                        "TITLE_CODE, " & _
                                        "GRADE_NAME, " & _
                                        "TITLE_NAME, " & _
                                        "USER_ORDER, " & _
                                        "EMPLOYEE_ID, " & _
                                        "SYSMAIL, " & _
                                        "CREATE_DATE" & _
                                        ")" & _
                                        " VALUES(" & strValues & ", '" & Date.Today & "')"
        End If

        Debug.Print(InsertSQL)

        SQLComm.CommandText = InsertSQL

        Try
            SQLComm.ExecuteNonQuery()
        Catch ex As Exception
            Console.WriteLine("ERROR - {0} Insert :" & ex.Message, i)
        End Try


    End Sub

    Private Sub DeleteUserInfo(ByVal CADSqlConn As SqlConnection, ByVal SQLComm As SqlCommand, Row As DataRow)

        Dim strUser_ID As String = Row.Field(Of String)("USER_ID")

        Dim DeleteSQL As String = "DELETE FROM INTERFACE_USERINFO WHERE USER_ID='" & strUser_ID & "'"

        Debug.Print(DeleteSQL)

        SQLComm.CommandText = DeleteSQL

        Try
            SQLComm.ExecuteNonQuery()
        Catch ex As Exception
            Console.WriteLine("ERROR - INTERFACE_USERINFO Delete :" & ex.Message)
        End Try

    End Sub

    Private Sub UpdateUesrInfo(ByVal CADSqlConn As SqlConnection, ByVal SQLComm As SqlCommand, Row As DataRow)

        Dim strUser_ID As String = Row.Field(Of String)("USER_ID")

        Dim UpdateSQL As String = String.Empty
        UpdateSQL = "UPDATE INTERFACE_USERINFO SET " & _
                                        "USER_NAME = '" & Row.Field(Of String)("USER_NAME") & "', " & _
                                        "USER_OTHER_NAME = '" & Row.Field(Of String)("USER_OTHER_NAME") & "', " & _
                                        "USER_DISPLAY_NAME = '" & Row.Field(Of String)("USER_DISPLAY_NAME") & "', " & _
                                        "USER_UID = '" & Row.Field(Of String)("USER_UID") & "', " & _
                                        "COMP_ID = '" & Row.Field(Of String)("COMP_ID") & "', " & _
                                        "COMP_NAME = '" & Row.Field(Of String)("COMP_NAME") & "', " & _
                                        "DEPT_ID = '" & Row.Field(Of String)("DEPT_ID") & "', " & _
                                        "DEPT_NAME = '" & Row.Field(Of String)("DEPT_NAME") & "', " & _
                                        "BIZ_CD = '" & Row.Field(Of String)("BIZ_CD") & "', " & _
                                        "M_ORG_NM = '" & Row.Field(Of String)("M_ORG_NM") & "', " & _
                                        "C_ORG_CD = '" & Row.Field(Of String)("C_ORG_CD") & "', " & _
                                        "C_ORG_NM = '" & Row.Field(Of String)("C_ORG_NM") & "', " & _
                                        "ORG_CD = '" & Row.Field(Of String)("ORG_CD") & "', " & _
                                        "ORGF_NM = '" & Row.Field(Of String)("ORGF_NM") & "', " & _
                                        "ORG_DISPLAY_NAME = '" & Row.Field(Of String)("ORG_DISPLAY_NAME") & "', " & _
                                        "OU_GROUP = '" & Row.Field(Of String)("OU_GROUP") & "', " & _
                                        "GRADE_CODE = '" & Row.Field(Of String)("GRADE_CODE") & "', " & _
                                        "TITLE_CODE = '" & Row.Field(Of String)("TITLE_CODE") & "', " & _
                                        "GRADE_NAME = '" & Row.Field(Of String)("GRADE_NAME") & "', " & _
                                        "TITLE_NAME = '" & Row.Field(Of String)("TITLE_NAME") & "', " & _
                                        "USER_ORDER = '" & Row.Field(Of String)("USER_ORDER") & "', " & _
                                        "EMPLOYEE_ID = '" & Row.Field(Of String)("EMPLOYEE_ID") & "', " & _
                                        "SYSMAIL = '" & Row.Field(Of String)("SYSMAIL") & "'" & _
                                        " WHERE USER_ID= '" & strUser_ID & "'"


        Debug.Print(UpdateSQL)

        SQLComm.CommandText = UpdateSQL

        Try
            SQLComm.ExecuteNonQuery()
        Catch ex As Exception
            Console.WriteLine("ERROR - INTERFACE_USERINFO Update :" & ex.Message)
        End Try

    End Sub

    Public Sub Interface_Organization(ByVal CADSqlConn As SqlConnection, ByVal SQLComm As SqlCommand, ModifyDate As String)

        Dim EPOraConn As New OracleConnection
        EPOraConn = _EP_Connect()
        Dim ORACmd As New OracleCommand

        Dim strSel As String = "SELECT ORG_NAME, " & _
                                        "ORG_OTHER_NAME, " & _
                                        "ORG_ABBR_NAME, " & _
                                        "ORG_ID, " & _
                                        "ORG_CODE, " & _
                                        "ORG_PARENT_ID, " & _
                                        "ORG_ORDER, " & _
                                        "SERVERS, " & _
                                        "ORG_TYPE, " & _
                                        "DESCRIPTION, " & _
                                        "COMPANY_ID " & _
                                        "FROM INTERFACE.V_ORGANIZATIONINFO_ALL " & _
                                            "WHERE(is_deleted = 0) ORDER BY ORG_NAME"
        Debug.Print(strSel)

        Dim ORAApt As New OracleDataAdapter(strSel, EPOraConn)
        Dim ORGInfo_DT As New DataTable
        ORAApt.Fill(ORGInfo_DT)
        ORAApt.Dispose()

        ORACmd.Dispose()

        EPOraConn.Close()
        EPOraConn.Dispose()
        EPOraConn = Nothing

        Dim DeleteSQL As String = "DELETE FROM INTERFACE_ORGANIZATIONINFO"

        Debug.Print(DeleteSQL)

        SQLComm.CommandText = DeleteSQL

        Try
            SQLComm.ExecuteNonQuery()
        Catch ex As Exception
            Console.WriteLine("ERROR - INTERFACE_USERINFO Delete :" & ex.Message)
        End Try


        For Each Row In ORGInfo_DT.AsEnumerable()

            Dim strValues As String = String.Empty
            For index = 0 To Row.ItemArray.Count - 1
                If index = 0 Then
                    strValues = "'" & Row.ItemArray(index) & "'"
                Else
                    strValues = strValues & ", '" & Row.ItemArray(index) & "'"
                End If
            Next

            Dim InsertSQL As String = String.Empty

            'InsertSQL = "INSERT INTO INTERFACE_ORGANIZATIONINFO(" & _
            '                            "ORG_NAME, " & _
            '                            "ORG_OTHER_NAME, " & _
            '                            "ORG_ABBR_NAME, " & _
            '                            "ORG_ID, " & _
            '                            "ORG_CODE, " & _
            '                            "ORG_PARENT_ID, " & _
            '                            "ORG_ORDER, " & _
            '                            "SERVERS, " & _
            '                            "ORG_TYPE, " & _
            '                            "DESCRIPTION, " & _
            '                            "COMPANY_ID " & _
            '                            ")" & _
            '                            " VALUES(" & strValues & ")"

            InsertSQL = String.Format("INSERT INTO INTERFACE_ORGANIZATIONINFO(" & _
                                        "ORG_NAME, " & _
                                        "ORG_OTHER_NAME, " & _
                                        "ORG_ABBR_NAME, " & _
                                        "ORG_ID, " & _
                                        "ORG_CODE, " & _
                                        "ORG_PARENT_ID, " & _
                                        "ORG_ORDER, " & _
                                        "SERVERS, " & _
                                        "ORG_TYPE, " & _
                                        "DESCRIPTION, " & _
                                        "COMPANY_ID, " & _
                                        "MODIFY_DATE" & _
                                        ")" & _
                                        " VALUES('{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}','{10}','{11}')",
                                        Row.Item("ORG_NAME"),
                                        Row.Item("ORG_OTHER_NAME"),
                                        Row.Item("ORG_ABBR_NAME"),
                                        Row.Item("ORG_ID"),
                                        Row.Item("ORG_CODE"),
                                        Row.Item("ORG_PARENT_ID"),
                                        Row.Item("ORG_ORDER"),
                                        Row.Item("SERVERS"),
                                        Row.Item("ORG_TYPE"),
                                        Row.Item("DESCRIPTION"),
                                        Row.Item("COMPANY_ID"),
                                        ModifyDate)
            Debug.Print(InsertSQL)

            SQLComm.CommandText = InsertSQL

            Try
                SQLComm.ExecuteNonQuery()
            Catch ex As Exception
                Console.WriteLine("ERROR - INTERFACE_ORGANIZATIONINFO Insert :" & ex.Message)
            End Try
        Next


    End Sub

#End Region



End Class

