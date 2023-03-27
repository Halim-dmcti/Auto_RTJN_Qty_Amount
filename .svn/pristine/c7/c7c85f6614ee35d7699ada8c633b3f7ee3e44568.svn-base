Imports System.Text
Public Class ClsGeneral
    Public Shared TANGGAL1, TANGGAL2 As DateTime
    Public Shared TAHUN, BULAN, TGL As Integer
    Public Shared NamaTable_Monitoring, jenis_mail_sender_report, jenis_mail_sender_report1, jenis_mail_sender_report2 As String

    Public Shared Function get_last_date(ByVal tanggal As DateTime) As Date
        Dim month_int As Integer = Month(DateAdd("m", 1, tanggal)) ' bulan + 1
        Dim year_int As Integer = Year(DateAdd("m", 1, tanggal))
        Dim Date_result As Date = DateSerial(year_int, month_int, 1) ' setting jadi tanggal 1 awal bulan berikutnya
        'get_last_date = "9/21/2021 12:00:00 AM"
        get_last_date = DateAdd("d", -1, Date_result) 'dikurangi 1 hari
    End Function


    Public Function get_kolom_HHmm(ByRef current_HHmm As Integer) As String
        Dim kolom_HHmm As String = ""

        'dimulai dari jam 08.30, karena periode jam DTI dari jam 07.30, sehingga baru mulai kalkulasi laporan awal jam 08.30
        'karena jam 07.30 masih sebagai kalkulasi laporan hari kemarin
        Select Case current_HHmm
            Case 830 To 929 'No.1 dari jam 08.30 s/d 09.29 WIB
                kolom_HHmm = "Pukul_08_30"

            Case 930 To 1029 'No.2 dari jam 09.30 s/d 10.29 WIB
                kolom_HHmm = "Pukul_09_30"

            Case 1030 To 1129 'No.3 dari jam 10.30 s/d 11.29 WIB
                kolom_HHmm = "Pukul_10_30"

            Case 1130 To 1229 'No.4 dari jam 11.30 s/d 12.29 WIB
                kolom_HHmm = "Pukul_11_30"

            Case 1230 To 1329 'No.5 dari jam 12.30 s/d 13.29 WIB
                kolom_HHmm = "Pukul_12_30"

            Case 1330 To 1429 'No.6 dari jam 13.30 s/d 14.29 WIB
                kolom_HHmm = "Pukul_13_30"

            Case 1430 To 1529 'No.7 dari jam 14.30 s/d 15.29 WIB
                kolom_HHmm = "Pukul_14_30"

            Case 1530 To 1629 'No.8 dari jam 15.30 s/d 16.29 WIB
                kolom_HHmm = "Pukul_15_30"

            Case 1630 To 1729 'No.9 dari jam 16.30 s/d 17.29 WIB
                kolom_HHmm = "Pukul_16_30"

            Case 1730 To 1829 'No.10 dari jam 17.30 s/d 18.29 WIB
                kolom_HHmm = "Pukul_17_30"

            Case 1830 To 1929 'No.11 dari jam 18.30 s/d 19.29 WIB
                kolom_HHmm = "Pukul_18_30"

            Case 1930 To 2029 'No.12 dari jam 19.30 s/d 20.29 WIB
                kolom_HHmm = "Pukul_19_30"

            Case 2030 To 2129 'No.13 dari jam 20.30 s/d 21.29 WIB
                kolom_HHmm = "Pukul_20_30"

            Case 2130 To 2229 'No.14 dari jam 21.30 s/d 22.29 WIB
                kolom_HHmm = "Pukul_21_30"

            Case 2230 To 2329 'No.15 dari jam 22.30 s/d 23.29 WIB
                kolom_HHmm = "Pukul_22_30"

            Case 2330 To 2359 'No.16 dari jam 23.30 s/d 23.59 WIB 'karena menggunakan range integer, sehingga range di pecah 2
                kolom_HHmm = "Pukul_23_30"
            Case 0 To 29 'No.16 dari jam 00.00 s/d 00.29 WIB 'karena menggunakan range integer, sehingga range di pecah 2
                kolom_HHmm = "Pukul_23_30"

            Case 30 To 129 'No.17 dari jam 00.30 s/d 01.29 WIB 
                kolom_HHmm = "Pukul_00_30"

            Case 130 To 229 'No.18 dari jam 01.30 s/d 02.29 WIB
                kolom_HHmm = "Pukul_01_30"

            Case 230 To 329 'No.19 dari jam 02.30 s/d 03.29 WIB
                kolom_HHmm = "Pukul_02_30"

            Case 330 To 429 'No.20 dari jam 03.30 s/d 04.29 WIB
                kolom_HHmm = "Pukul_03_30"

            Case 430 To 529 'No.21 dari jam 04.30 s/d 05.29 WIB
                kolom_HHmm = "Pukul_04_30"

            Case 530 To 629 'No.22 dari jam 05.30 s/d 06.29 WIB
                kolom_HHmm = "Pukul_05_30"

            Case 630 To 729 'No.23 dari jam 06.30 s/d 07.29 WIB
                kolom_HHmm = "Pukul_06_30"

            Case 730 To 829 'No.24 dari jam 07.30 s/d 08.29 WIB
                kolom_HHmm = "Pukul_07_30"

            Case Else 'status email monitoring mail sender
                kolom_HHmm = "EmailMonitoring"

        End Select

        get_kolom_HHmm = kolom_HHmm
    End Function


    Public Function cek_status_sudah_email(
                                            ByVal currentDate As Date,
                                            ByRef current_HHmm As Integer,
                                            ByVal JenisLaporan As String) As Boolean
        Dim hasil_cek As Boolean = False
        Dim query As New StringBuilder
        Dim dt As DataTable
        Dim kolom_HHmm As String = get_kolom_HHmm(current_HHmm)

        If JenisLaporan = "qty_amount" Then
            JenisLaporan = "Realtime Production (Amount)"
        Else
            JenisLaporan = "Realtime Production (Qty)"
        End If

        query.AppendLine(" select ")
        query.AppendLine("     date ")
        query.AppendLine("     ," & kolom_HHmm & " ")
        query.AppendLine("     ,last_email_sent ")
        query.AppendLine(" from ")
        query.AppendLine("     ad_dis_monitoring_maintenance ")
        query.AppendLine(" where ")
        query.AppendLine("     jenis_mail_sender LIKE '%" & JenisLaporan & "%' ")
        query.AppendLine("     and date = '" & currentDate.ToString("yyyy-MM-dd") & "' ")
        query.AppendLine("     and " & kolom_HHmm & " = 'OK' ")
        dt = ClsConfig.ExecuteQuery(query.ToString(), ClsConfig.IPServer_ADDONS)
        query.Length = 0
        query.Capacity = 0

        If dt.Rows.Count > 0 Then
            hasil_cek = True
        End If

        cek_status_sudah_email = hasil_cek
    End Function

    Public Shared Function aktual_produksi(ByVal current_date As DateTime)

        Dim result_string As String = "0"

        Dim query As StringBuilder = New StringBuilder()
        query.AppendLine(" select ")
        query.AppendLine("     	top(1) CASE WHEN id_seisan IS NULL THEN 0 ELSE 1 END AS aktual_produksi")
        query.AppendLine(" from ")
        query.AppendLine("     Z_RT_data_J_kotei ")
        query.AppendLine(" where ")
        query.AppendLine("     id_kotei = '5230'                                               ")
        query.AppendLine("     AND shift_date = '" & current_date.ToString("yyyyMMdd") & "' ")
        query.AppendLine(" ")

        Dim dt As New System.Data.DataTable
        dt = ClsConfig.ExecuteQuery(query.ToString(), ClsConfig.IPServer_RTJN_PRD)
        query.Length = 0
        query.Capacity = 0
        If dt.Rows.Count > 0 Then
            result_string = dt(0)("aktual_produksi").ToString()
        Else
            result_string = "0"
        End If
        Return result_string

    End Function

    Public Shared Function target_qty_amount(ByVal current_date As DateTime)

        Dim result_string As String = "0"

        Dim query As StringBuilder = New StringBuilder()
        query.AppendLine(" select ")
        query.AppendLine("     	top(1) CASE WHEN target_qty_jam_ke_1 = 0 THEN 0 ELSE 1 END AS target_qty")
        query.AppendLine(" from ")
        query.AppendLine("     ad_dis_rtjn_sum_qty_amount_target ")
        query.AppendLine(" where ")
        query.AppendLine("     target_date = '" & current_date.ToString("yyyy-MM-dd") & "' ")
        query.AppendLine(" ")

        Dim dt As New System.Data.DataTable
        dt = ClsConfig.ExecuteQuery(query.ToString(), ClsConfig.IPServer_ADDONS)
        query.Length = 0
        query.Capacity = 0
        If dt.Rows.Count > 0 Then
            result_string = dt(0)("target_qty").ToString()
        Else
            result_string = "0"
        End If
        Return result_string

    End Function


    Public Shared Function find_target_print_qty(ByVal target_date As Date) As String
        Dim result_string As String = ""

        Dim query As StringBuilder = New StringBuilder()
        query.AppendLine(" select ")
        query.AppendLine("     target_date ")
        query.AppendLine("     ,isnull(target_print_qty,0) target_print_qty ")
        query.AppendLine(" from ")
        query.AppendLine("     ad_dis_rtjn_sum_qty_amount_target ")
        query.AppendLine(" where ")
        query.AppendLine("     target_date = '" & target_date.ToString("yyyy-MM-dd") & "' ")
        query.AppendLine(" ")

        Dim dt As New System.Data.DataTable
        dt = ClsConfig.ExecuteQuery(query.ToString(), ClsConfig.IPServer_ADDONS)
        query.Length = 0
        query.Capacity = 0
        If dt.Rows.Count > 0 Then
            result_string = dt(0)("target_print_qty").ToString()
        End If
        Return result_string
    End Function

    Public Shared Function find_target_qty(ByVal target_date As Date) As String
        Dim result_string As String = ""

        Dim query As StringBuilder = New StringBuilder()
        query.AppendLine(" select ")
        query.AppendLine("     target_date ")
        query.AppendLine("     ,(isnull(target_qty_jam_ke_1,0) + ")
        query.AppendLine("     isnull(target_qty_jam_ke_2,0) + ")
        query.AppendLine("     isnull(target_qty_jam_ke_3,0) + ")
        query.AppendLine("     isnull(target_qty_jam_ke_4,0) + ")
        query.AppendLine("     isnull(target_qty_jam_ke_5,0) + ")
        query.AppendLine("     isnull(target_qty_jam_ke_6,0) + ")
        query.AppendLine("     isnull(target_qty_jam_ke_7,0) + ")
        query.AppendLine("     isnull(target_qty_jam_ke_8,0) + ")
        query.AppendLine("     isnull(target_qty_jam_ke_9,0) + ")
        query.AppendLine("     isnull(target_qty_jam_ke_10,0) + ")
        query.AppendLine("     isnull(target_qty_jam_ke_11,0) + ")
        query.AppendLine("     isnull(target_qty_jam_ke_12,0) + ")
        query.AppendLine("     isnull(target_qty_jam_ke_13,0) + ")
        query.AppendLine("     isnull(target_qty_jam_ke_14,0) + ")
        query.AppendLine("     isnull(target_qty_jam_ke_17,0) + ")
        query.AppendLine("     isnull(target_qty_jam_ke_18,0) + ")
        query.AppendLine("     isnull(target_qty_jam_ke_19,0) + ")
        query.AppendLine("     isnull(target_qty_jam_ke_20,0) + ")
        query.AppendLine("     isnull(target_qty_jam_ke_21,0) + ")
        query.AppendLine("     isnull(target_qty_jam_ke_22,0)) target_total_qty_per_jam ")
        query.AppendLine(" from ")
        query.AppendLine("     ad_dis_rtjn_sum_qty_amount_target ")
        query.AppendLine(" where ")
        query.AppendLine("     target_date = '" & target_date.ToString("yyyy-MM-dd") & "' ")
        query.AppendLine(" ")

        Dim dt As New System.Data.DataTable
        dt = ClsConfig.ExecuteQuery(query.ToString(), ClsConfig.IPServer_ADDONS)
        query.Length = 0
        query.Capacity = 0
        If dt.Rows.Count > 0 Then
            result_string = dt(0)("target_total_qty_per_jam").ToString()
        End If
        Return result_string
    End Function

    Public Shared Function find_target_amount(ByVal target_date As Date) As String
        Dim result_string As String = ""

        Dim query As StringBuilder = New StringBuilder()
        query.AppendLine(" select ")
        query.AppendLine("     target_date ")
        query.AppendLine("     ,(isnull(target_amount_jam_ke_1,0) + ")
        query.AppendLine("     isnull(target_amount_jam_ke_2,0) + ")
        query.AppendLine("     isnull(target_amount_jam_ke_3,0) + ")
        query.AppendLine("     isnull(target_amount_jam_ke_4,0) + ")
        query.AppendLine("     isnull(target_amount_jam_ke_5,0) + ")
        query.AppendLine("     isnull(target_amount_jam_ke_6,0) + ")
        query.AppendLine("     isnull(target_amount_jam_ke_7,0) + ")
        query.AppendLine("     isnull(target_amount_jam_ke_8,0) + ")
        query.AppendLine("     isnull(target_amount_jam_ke_9,0) + ")
        query.AppendLine("     isnull(target_amount_jam_ke_10,0) + ")
        query.AppendLine("     isnull(target_amount_jam_ke_11,0) + ")
        query.AppendLine("     isnull(target_amount_jam_ke_12,0) + ")
        query.AppendLine("     isnull(target_amount_jam_ke_13,0) + ")
        query.AppendLine("     isnull(target_amount_jam_ke_14,0) + ")
        query.AppendLine("     isnull(target_amount_jam_ke_17,0) + ")
        query.AppendLine("     isnull(target_amount_jam_ke_18,0) + ")
        query.AppendLine("     isnull(target_amount_jam_ke_19,0) + ")
        query.AppendLine("     isnull(target_amount_jam_ke_20,0) + ")
        query.AppendLine("     isnull(target_amount_jam_ke_21,0) + ")
        query.AppendLine("     isnull(target_amount_jam_ke_22,0)) target_total_amount_per_jam ")
        query.AppendLine(" from ")
        query.AppendLine("     ad_dis_rtjn_sum_qty_amount_target ")
        query.AppendLine(" where ")
        query.AppendLine("     target_date = '" & target_date.ToString("yyyy-MM-dd") & "' ")
        query.AppendLine(" ")

        Dim dt As New System.Data.DataTable
        dt = ClsConfig.ExecuteQuery(query.ToString(), ClsConfig.IPServer_ADDONS)
        query.Length = 0
        query.Capacity = 0
        If dt.Rows.Count > 0 Then
            result_string = dt(0)("target_total_amount_per_jam").ToString()
        End If
        Return result_string
    End Function

    Public Shared Function find_shift_group(ByVal target_date As Date, ByVal shift As Int32) As String
        Dim result_string As String = ""
        Dim query As StringBuilder = New StringBuilder()
        query.AppendLine(" select ")
        query.AppendLine("     shift_date, shift, grp ")
        query.AppendLine(" from ")
        query.AppendLine("     Z_RT_data_group ")
        query.AppendLine(" where ")
        query.AppendLine("     shift_date = '" & target_date.ToString("yyyy-MM-dd") & "' ")
        query.AppendLine("     and shift = " & shift.ToString() & " ")
        query.AppendLine(" ")

        Dim dt As New System.Data.DataTable
        dt = ClsConfig.ExecuteQuery(query.ToString(), ClsConfig.IPServer_RTJN_PRD)
        query.Length = 0
        query.Capacity = 0
        If dt.Rows.Count > 0 Then
            result_string = dt(0)("grp").ToString()
        End If
        Return result_string
    End Function

    Public Sub monitoring_email(ByVal currentDate As Date, ByRef current_HHmm As Integer, ByVal jenis_laporan As String, ByVal message As String)

        Try
            'Table Monitoring Maintenance
            NamaTable_Monitoring = "ad_dis_monitoring_maintenance"

            'Default Jenis Laporan
            jenis_mail_sender_report = "Realtime"

            If (jenis_laporan = "qty") Then
                jenis_mail_sender_report = "Realtime Production (Qty)"
            ElseIf (jenis_laporan = "qty_amount") Then
                jenis_mail_sender_report = "Realtime Production (Amount)"
            End If

            'Dim judgementResult As String
            'Dim masalah_kegagalan As String
            'Dim jumlah_kegagalan As Integer

            'If (message = "") Then
            '    judgementResult = "OK"
            '    masalah_kegagalan = ""
            '    jumlah_kegagalan = 0
            'ElseIf (message = "Tidak ada transaksi") Then
            '    judgementResult = "-"
            '    masalah_kegagalan = "Tidak ada transaksi pada proses Gaikan 2x di aktual Produksi."
            '    jumlah_kegagalan = 0
            'Else
            '    judgementResult = "NG"
            '    masalah_kegagalan = message.Replace("'", ControlChars.Quote)
            '    jumlah_kegagalan = 1
            'End If

            Dim query As New StringBuilder
            Dim dt As DataTable
            Dim kolom_HHmm As String = get_kolom_HHmm(current_HHmm)

            query.AppendLine(" select ")
            query.AppendLine("     date ")
            query.AppendLine("     ,last_email_sent ")
            query.AppendLine("     ,masalah_kegagalan ")
            query.AppendLine(" from ")
            query.AppendLine("     " & NamaTable_Monitoring & " ")
            query.AppendLine(" where ")
            query.AppendLine("     jenis_mail_sender like '%" & jenis_mail_sender_report & "%' ")
            query.AppendLine("     and date = '" & currentDate.ToString("yyyy-MM-dd") & "' ")
            dt = ClsConfig.ExecuteQuery(query.ToString(), ClsConfig.IPServer_ADDONS)
            query.Length = 0
            query.Capacity = 0

            If dt.Rows.Count > 0 Then
                query.AppendLine(" update ")
                query.AppendLine("     " & NamaTable_Monitoring & " ")
                query.AppendLine(" set ")
                query.AppendLine("     " & kolom_HHmm & " = 'OK' ")
                query.AppendLine("     ,jumlah_kegagalan = ( IIF(Pukul_08_30 = 'OK', 0, 1) + IIF(Pukul_09_30 = 'OK', 0, 1) + IIF(Pukul_10_30 = 'OK', 0, 1) + ")
                query.AppendLine("                           IIF(Pukul_11_30 = 'OK', 0, 1) + IIF(Pukul_12_30 = 'OK', 0, 1) + IIF(Pukul_13_30 = 'OK', 0, 1) + ")
                query.AppendLine("                           IIF(Pukul_14_30 = 'OK', 0, 1) + IIF(Pukul_15_30 = 'OK', 0, 1) + IIF(Pukul_16_30 = 'OK', 0, 1) + ")
                query.AppendLine("                           IIF(Pukul_17_30 = 'OK', 0, 1) + IIF(Pukul_18_30 = 'OK', 0, 1) + IIF(Pukul_19_30 = 'OK', 0, 1) + ")
                query.AppendLine("                           IIF(Pukul_20_30 = 'OK', 0, 1) + IIF(Pukul_21_30 = 'OK', 0, 1) + IIF(Pukul_22_30 = 'OK', 0, 1) + ")
                query.AppendLine("                           IIF(Pukul_23_30 = 'OK', 0, 1) + IIF(Pukul_00_30 = 'OK', 0, 1) + IIF(Pukul_01_30 = 'OK', 0, 1) + ")
                query.AppendLine("                           IIF(Pukul_02_30 = 'OK', 0, 1) + IIF(Pukul_03_30 = 'OK', 0, 1) + IIF(Pukul_04_30 = 'OK', 0, 1) + ")
                query.AppendLine("                           IIF(Pukul_05_30 = 'OK', 0, 1) + IIF(Pukul_06_30 = 'OK', 0, 1) + IIF(Pukul_07_30 = 'OK', 0, 1) ) ")
                query.AppendLine("     ,last_email_sent = getdate() ")
                query.AppendLine(" where ")
                query.AppendLine("     jenis_mail_sender = '" & jenis_mail_sender_report & "' ")
                query.AppendLine("     and date = '" & currentDate.ToString("yyyy-MM-dd") & "' ")
                ClsConfig.ExecuteNonQuery(query.ToString(), ClsConfig.IPServer_ADDONS)
                query.Length = 0
                query.Capacity = 0

                Console.WriteLine("--> Proses Update Data File Monitoring Maintenance: " & jenis_mail_sender_report)
            Else

                query.AppendLine(" Insert Into ")
                query.AppendLine("     " & NamaTable_Monitoring & " ")
                query.AppendLine("     ( ")
                query.AppendLine("         date ")
                query.AppendLine("         ,jenis_mail_sender ")
                query.AppendLine("         ," & kolom_HHmm & " ")
                query.AppendLine("         ,jumlah_kegagalan ")
                query.AppendLine("         ,last_email_sent ")
                query.AppendLine("     ) ")
                query.AppendLine(" values ")
                query.AppendLine("     ( ")
                query.AppendLine("         '" & currentDate.ToString("yyyy-MM-dd") & "' ")
                query.AppendLine("         ,'" & jenis_mail_sender_report & "' ")
                query.AppendLine("         ,'OK' ")
                query.AppendLine("         , 0 ")
                query.AppendLine("         ,getdate() ")
                query.AppendLine("     ) ")
                ClsConfig.ExecuteNonQuery(query.ToString(), ClsConfig.IPServer_ADDONS)
                query.Length = 0
                query.Capacity = 0

                Console.WriteLine("--> Proses Input Data File Monitoring Maintenance: " & jenis_mail_sender_report)
            End If

        Catch ex As Exception
            monitoring_email(currentDate, current_HHmm, jenis_laporan, message)

            ClsConfig.create_log_error(currentDate, jenis_laporan, "[" + DateTime.Now.ToString("dd-MM-yyyy hh:mm:ss") + "] -- [ " + ex.Message + " ] -- Proses insert/update monitoring Realtime Production error")
            Environment.Exit(0)
        End Try
    End Sub

End Class
