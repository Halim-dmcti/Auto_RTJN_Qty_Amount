﻿Imports Microsoft.VisualBasic
Imports System.Data.SqlClient
Imports System.Text
Imports System.Reflection
Imports System.Net.Mail
Imports System
Imports System.Globalization

Module Module1

    Dim SQL As String
    Dim ClsGnrl As New ClsGeneral
    Dim current_HHmm As Integer 'jam dan menit yang sedang berjalan
    Dim ClsAutRep As New ClsAutoReport
    Dim start_date, end_date, current_date As DateTime
    Dim myCulture As System.Globalization.CultureInfo = Globalization.CultureInfo.CurrentCulture
    Dim TAHUN, BULAN, TGL As Integer

    'jenis laporan email
    Dim jenis_laporan As String
    Dim jenis_mail_sender_report As String


    Dim now_oclock As String
    Dim str_status_email_amount As String
    Dim str_status_email_qty As String
    Dim str_last_email_amount As String
    Dim str_last_email_qty As String
    Dim status_sudah_email As Boolean 'jika sudah email maka nilai true, jika belum nilai false
    Dim subject_email As String

    'cek target
    Dim sts_target As String
    Dim sts_aktual_produksi As String

    Sub Main()

        Console.WriteLine("PROGRAM GET QTY AND AMOUNT")
        Console.WriteLine("")

        Console.WriteLine("##START SET CONFIG")
        ClsConfig.get_variable_setting()
        Console.WriteLine("##FINISH SET CONFIG")
        Console.WriteLine("")

        current_date = Now
        'current_date = Now.AddDays(-1)
        'current_date = Now.AddHours(1)

        'Jika < jam 08.29 WIB masih menggunakan periode tanggal kemarin (menggunakan jam 08.29 karena proses kalkulasi pasti ada delay dan pasti > 07.30 WIB)
        'menggunakan 08.29 WIB karena kalkulasi pertama per-hari dijam 08.30
        'fungsi ini untuk mengikuti periode hari produksi DTI 07.30 s/d 07.30 dan juga handle periode pada irisan bulan
        '(akhir bulan dengan awal bulan contoh : tgl. 01-11-22 jam 07.00 dihitung masih bulan sebelumnya)
        If Int32.Parse(current_date.ToString("HHmm")) <= 829 Then
            current_date = current_date.AddDays(-1)
        End If

        start_date = DateSerial(Year(current_date), Month(current_date), 1)
        end_date = ClsGeneral.get_last_date(current_date)
        current_HHmm = Int32.Parse(current_date.ToString("HHmm"))

        Dim kolom_HHmm As String = ClsGnrl.get_kolom_HHmm(current_HHmm)

        str_status_email_amount = ClsGnrl.cek_status_sudah_email(current_date, current_HHmm, "qty_amount")
        str_status_email_qty = ClsGnrl.cek_status_sudah_email(current_date, current_HHmm, "qty")

        'Cek Target (qty_amount_target)
        sts_target = ClsGeneral.target_qty_amount(current_date)

        'Cek Aktual Produksi
        sts_aktual_produksi = ClsGeneral.aktual_produksi(current_date)

        Dim message As String = ""

        now_oclock = current_date.ToString("HH")
        If (str_status_email_amount = True And str_status_email_qty = True) Then
            'JIKA PROGRAM BERHASIL, DI-SKIP PROSES
            Console.WriteLine("PROGRAM SUDAH KALKULASI DAN KIRIM EMAIL")
        Else
            get_calculation(start_date, end_date, current_date, current_HHmm)
        End If

        'kirim laporan monitoring mail sender qty & amount hanya jam 07.30
        If Int32.Parse(current_date.ToString("HHmm")) >= 730 And Int32.Parse(current_date.ToString("HHmm")) <= 829 Then
            'kirim laporan monitoring amount
            str_status_email_amount = ClsGnrl.cek_status_sudah_email(current_date, 9999, "qty_amount")
            If str_status_email_amount = True Then
                Console.WriteLine("PROGRAM SUDAH KIRIM EMAIL MONITORING QTY AMOUNT")
            Else
                ClsAutRep.AutoReportMonitoring(start_date, current_date, "qty_amount")
            End If

            'krim laporan monitoring qty
            str_status_email_qty = ClsGnrl.cek_status_sudah_email(current_date, 9999, "qty")
            If str_status_email_qty = True Then
                Console.WriteLine("PROGRAM SUDAH KIRIM EMAIL MONITORING QTY")
            Else
                ClsAutRep.AutoReportMonitoring(start_date, current_date, "qty")
            End If
        End If

    End Sub

    Private Sub get_calculation(ByVal start_date As Date,
                                ByVal end_date As Date,
                                ByVal currentDate As Date,
                                ByVal currentHHmm As Integer)

        Dim dt As DataTable
        'current_date = Now

        ''kondisi ini jika ada irisan data antara akhir bulan dengan awal bulan
        ''contoh = pada tgl. 1 agustus jam 07.30, masih periode bulan juli karena masih jam shift 2 (berdasarkan jam RTJN atau jam kerja DMCTI)
        ''dibuat < jam 08.00 untuk antisipasi jika terjadi delay proses program dari server/ task scheduler
        'If current_date.ToString("dd") = "01" And Int32.Parse(current_date.ToString("HHmm")) < 1700 Then
        '    current_date = current_date.AddDays(-1)
        'End If

        'dibuat < jam 08.00 untuk antisipasi jika terjadi delay proses program dari server/ task scheduler
        'If Int32.Parse(current_date.ToString("HHmm")) < 800 Then
        '    current_date = current_date.AddDays(-1)
        'End If

        TAHUN = Year(current_date)
        BULAN = Month(current_date)
        TGL = Day(current_date)

        jenis_laporan = "all"


        Console.WriteLine("##PROSES GET DATA FROM DATABASE")

        Dim query As StringBuilder = New StringBuilder()
        query.AppendLine(" select ")
        query.AppendLine(" 	DATE ")
        query.AppendLine(" 	,sum(QTY_OK) QTY_OK ")
        query.AppendLine(" 	,sum(AMOUNT) AMOUNT ")
        query.AppendLine(" 	,sum(TOTAL_SIZE) TOTAL_SIZE ")
        query.AppendLine(" 	,sum(act_qty_jam_ke_1) act_qty_jam_ke_1 ")
        query.AppendLine(" 	,sum(act_amount_jam_ke_1) act_amount_jam_ke_1 ")
        query.AppendLine(" 	,sum(act_qty_jam_ke_2) act_qty_jam_ke_2 ")
        query.AppendLine(" 	,sum(act_amount_jam_ke_2) act_amount_jam_ke_2 ")
        query.AppendLine(" 	,sum(act_qty_jam_ke_3) act_qty_jam_ke_3 ")
        query.AppendLine(" 	,sum(act_amount_jam_ke_3) act_amount_jam_ke_3 ")
        query.AppendLine(" 	,sum(act_qty_jam_ke_4) act_qty_jam_ke_4 ")
        query.AppendLine(" 	,sum(act_amount_jam_ke_4) act_amount_jam_ke_4 ")
        query.AppendLine(" 	,sum(act_qty_jam_ke_5) act_qty_jam_ke_5 ")
        query.AppendLine(" 	,sum(act_amount_jam_ke_5) act_amount_jam_ke_5 ")
        query.AppendLine(" 	,sum(act_qty_jam_ke_6) act_qty_jam_ke_6 ")
        query.AppendLine(" 	,sum(act_amount_jam_ke_6) act_amount_jam_ke_6 ")
        query.AppendLine(" 	,sum(act_qty_jam_ke_7) act_qty_jam_ke_7 ")
        query.AppendLine(" 	,sum(act_amount_jam_ke_7) act_amount_jam_ke_7 ")
        query.AppendLine(" 	,sum(act_qty_jam_ke_8) act_qty_jam_ke_8 ")
        query.AppendLine(" 	,sum(act_amount_jam_ke_8) act_amount_jam_ke_8 ")
        query.AppendLine(" 	,sum(act_qty_jam_ke_9) act_qty_jam_ke_9 ")
        query.AppendLine(" 	,sum(act_amount_jam_ke_9) act_amount_jam_ke_9 ")
        query.AppendLine(" 	,sum(act_qty_jam_ke_10) act_qty_jam_ke_10 ")
        query.AppendLine(" 	,sum(act_amount_jam_ke_10) act_amount_jam_ke_10 ")
        query.AppendLine(" 	,sum(act_qty_jam_ke_11) act_qty_jam_ke_11 ")
        query.AppendLine(" 	,sum(act_amount_jam_ke_11) act_amount_jam_ke_11 ")
        query.AppendLine(" 	,sum(act_qty_jam_ke_12) act_qty_jam_ke_12 ")
        query.AppendLine(" 	,sum(act_amount_jam_ke_12) act_amount_jam_ke_12 ")
        query.AppendLine(" 	,sum(act_qty_jam_ke_13) act_qty_jam_ke_13 ")
        query.AppendLine(" 	,sum(act_amount_jam_ke_13) act_amount_jam_ke_13 ")
        query.AppendLine(" 	,sum(act_qty_jam_ke_14) act_qty_jam_ke_14 ")
        query.AppendLine(" 	,sum(act_amount_jam_ke_14) act_amount_jam_ke_14 ")
        query.AppendLine(" 	,sum(act_qty_jam_ke_15) act_qty_jam_ke_15 ")
        query.AppendLine(" 	,sum(act_amount_jam_ke_15) act_amount_jam_ke_15 ")
        query.AppendLine(" 	,sum(act_qty_jam_ke_16) act_qty_jam_ke_16 ")
        query.AppendLine(" 	,sum(act_amount_jam_ke_16) act_amount_jam_ke_16 ")
        query.AppendLine(" 	,sum(act_qty_jam_ke_17) act_qty_jam_ke_17 ")
        query.AppendLine(" 	,sum(act_amount_jam_ke_17) act_amount_jam_ke_17 ")
        query.AppendLine(" 	,sum(act_qty_jam_ke_18) act_qty_jam_ke_18 ")
        query.AppendLine(" 	,sum(act_amount_jam_ke_18) act_amount_jam_ke_18 ")
        query.AppendLine(" 	,sum(act_qty_jam_ke_19) act_qty_jam_ke_19 ")
        query.AppendLine(" 	,sum(act_amount_jam_ke_19) act_amount_jam_ke_19 ")
        query.AppendLine(" 	,sum(act_qty_jam_ke_20) act_qty_jam_ke_20 ")
        query.AppendLine(" 	,sum(act_amount_jam_ke_20) act_amount_jam_ke_20 ")
        query.AppendLine(" 	,sum(act_qty_jam_ke_21) act_qty_jam_ke_21 ")
        query.AppendLine(" 	,sum(act_amount_jam_ke_21) act_amount_jam_ke_21 ")
        query.AppendLine(" 	,sum(act_qty_jam_ke_22) act_qty_jam_ke_22 ")
        query.AppendLine(" 	,sum(act_amount_jam_ke_22) act_amount_jam_ke_22 ")
        query.AppendLine(" from ")
        query.AppendLine(" 	( ")
        query.AppendLine(" 		select ")
        query.AppendLine(" 			A.id_seihin DMC_CODE ")
        query.AppendLine(" 			,convert(date,A.shift_date) DATE ")
        query.AppendLine(" 			,A.amnt_OK QTY_OK ")
        query.AppendLine(" 			,(case when A.time_sagyo between CONVERT(varchar, convert(date,A.shift_date), 101) +' 07:30:00' and CONVERT(varchar, convert(date,A.shift_date), 101) +' 08:29:59' then A.amnt_OK else 0 end) act_qty_jam_ke_1 ")
        query.AppendLine(" 			,(case when A.time_sagyo between CONVERT(varchar, convert(date,A.shift_date), 101) +' 07:30:00' and CONVERT(varchar, convert(date,A.shift_date), 101) +' 08:29:59' then A.amnt_OK else 0 end) * isnull(D.price, 0) act_amount_jam_ke_1 ")
        query.AppendLine(" 			,(case when A.time_sagyo between CONVERT(varchar, convert(date,A.shift_date), 101) +' 08:30:00' and CONVERT(varchar, convert(date,A.shift_date), 101) +' 09:29:59' then A.amnt_OK else 0 end) act_qty_jam_ke_2 ")
        query.AppendLine(" 			,(case when A.time_sagyo between CONVERT(varchar, convert(date,A.shift_date), 101) +' 08:30:00' and CONVERT(varchar, convert(date,A.shift_date), 101) +' 09:29:59' then A.amnt_OK else 0 end) * isnull(D.price, 0) act_amount_jam_ke_2 ")
        query.AppendLine(" 			,(case when A.time_sagyo between CONVERT(varchar, convert(date,A.shift_date), 101) +' 09:30:00' and CONVERT(varchar, convert(date,A.shift_date), 101) +' 10:39:59' then A.amnt_OK else 0 end) act_qty_jam_ke_3 ")
        query.AppendLine(" 			,(case when A.time_sagyo between CONVERT(varchar, convert(date,A.shift_date), 101) +' 09:30:00' and CONVERT(varchar, convert(date,A.shift_date), 101) +' 10:39:59' then A.amnt_OK else 0 end) * isnull(D.price, 0) act_amount_jam_ke_3 ")
        query.AppendLine(" 			,(case when A.time_sagyo between CONVERT(varchar, convert(date,A.shift_date), 101) +' 10:40:00' and CONVERT(varchar, convert(date,A.shift_date), 101) +' 12:29:59' then A.amnt_OK else 0 end) act_qty_jam_ke_4 ")
        query.AppendLine(" 			,(case when A.time_sagyo between CONVERT(varchar, convert(date,A.shift_date), 101) +' 10:40:00' and CONVERT(varchar, convert(date,A.shift_date), 101) +' 12:29:59' then A.amnt_OK else 0 end) * isnull(D.price, 0) act_amount_jam_ke_4 ")
        query.AppendLine(" 			,(case when A.time_sagyo between CONVERT(varchar, convert(date,A.shift_date), 101) +' 12:30:00' and CONVERT(varchar, convert(date,A.shift_date), 101) +' 13:29:59' then A.amnt_OK else 0 end) act_qty_jam_ke_5 ")
        query.AppendLine(" 			,(case when A.time_sagyo between CONVERT(varchar, convert(date,A.shift_date), 101) +' 12:30:00' and CONVERT(varchar, convert(date,A.shift_date), 101) +' 13:29:59' then A.amnt_OK else 0 end) * isnull(D.price, 0) act_amount_jam_ke_5 ")
        query.AppendLine(" 			,(case when A.time_sagyo between CONVERT(varchar, convert(date,A.shift_date), 101) +' 13:30:00' and CONVERT(varchar, convert(date,A.shift_date), 101) +' 14:29:59' then A.amnt_OK else 0 end) act_qty_jam_ke_6 ")
        query.AppendLine(" 			,(case when A.time_sagyo between CONVERT(varchar, convert(date,A.shift_date), 101) +' 13:30:00' and CONVERT(varchar, convert(date,A.shift_date), 101) +' 14:29:59' then A.amnt_OK else 0 end) * isnull(D.price, 0) act_amount_jam_ke_6 ")
        query.AppendLine(" 			,(case when A.time_sagyo between CONVERT(varchar, convert(date,A.shift_date), 101) +' 14:30:00' and CONVERT(varchar, convert(date,A.shift_date), 101) +' 15:59:59' then A.amnt_OK else 0 end) act_qty_jam_ke_7 ")
        query.AppendLine(" 			,(case when A.time_sagyo between CONVERT(varchar, convert(date,A.shift_date), 101) +' 14:30:00' and CONVERT(varchar, convert(date,A.shift_date), 101) +' 15:59:59' then A.amnt_OK else 0 end) * isnull(D.price, 0) act_amount_jam_ke_7 ")
        query.AppendLine(" 			,(case when A.time_sagyo between CONVERT(varchar, convert(date,A.shift_date), 101) +' 16:00:00' and CONVERT(varchar, convert(date,A.shift_date), 101) +' 16:59:59' then A.amnt_OK else 0 end) act_qty_jam_ke_8 ")
        query.AppendLine(" 			,(case when A.time_sagyo between CONVERT(varchar, convert(date,A.shift_date), 101) +' 16:00:00' and CONVERT(varchar, convert(date,A.shift_date), 101) +' 16:59:59' then A.amnt_OK else 0 end) * isnull(D.price, 0) act_amount_jam_ke_8 ")
        query.AppendLine(" 			,(case when A.time_sagyo between CONVERT(varchar, convert(date,A.shift_date), 101) +' 17:00:00' and CONVERT(varchar, convert(date,A.shift_date), 101) +' 18:29:59' then A.amnt_OK else 0 end) act_qty_jam_ke_9 ")
        query.AppendLine(" 			,(case when A.time_sagyo between CONVERT(varchar, convert(date,A.shift_date), 101) +' 17:00:00' and CONVERT(varchar, convert(date,A.shift_date), 101) +' 18:29:59' then A.amnt_OK else 0 end) * isnull(D.price, 0) act_amount_jam_ke_9 ")
        query.AppendLine(" 			,(case when A.time_sagyo between CONVERT(varchar, convert(date,A.shift_date), 101) +' 18:30:00' and CONVERT(varchar, convert(date,A.shift_date), 101) +' 19:29:59' then A.amnt_OK else 0 end) act_qty_jam_ke_10 ")
        query.AppendLine(" 			,(case when A.time_sagyo between CONVERT(varchar, convert(date,A.shift_date), 101) +' 18:30:00' and CONVERT(varchar, convert(date,A.shift_date), 101) +' 19:29:59' then A.amnt_OK else 0 end) * isnull(D.price, 0) act_amount_jam_ke_10 ")
        query.AppendLine(" 			,(case when A.time_sagyo between CONVERT(varchar, convert(date,A.shift_date), 101) +' 19:30:00' and CONVERT(varchar, convert(date,A.shift_date), 101) +' 20:29:59' then A.amnt_OK else 0 end) act_qty_jam_ke_11 ")
        query.AppendLine(" 			,(case when A.time_sagyo between CONVERT(varchar, convert(date,A.shift_date), 101) +' 19:30:00' and CONVERT(varchar, convert(date,A.shift_date), 101) +' 20:29:59' then A.amnt_OK else 0 end) * isnull(D.price, 0) act_amount_jam_ke_11 ")
        query.AppendLine(" 			,(case when A.time_sagyo between CONVERT(varchar, convert(date,A.shift_date), 101) +' 20:30:00' and CONVERT(varchar, convert(date,A.shift_date), 101) +' 21:29:59' then A.amnt_OK else 0 end) act_qty_jam_ke_12 ")
        query.AppendLine(" 			,(case when A.time_sagyo between CONVERT(varchar, convert(date,A.shift_date), 101) +' 20:30:00' and CONVERT(varchar, convert(date,A.shift_date), 101) +' 21:29:59' then A.amnt_OK else 0 end) * isnull(D.price, 0) act_amount_jam_ke_12 ")
        query.AppendLine(" 			,(case when A.time_sagyo between CONVERT(varchar, convert(date,A.shift_date), 101) +' 21:30:00' and CONVERT(varchar, convert(date,A.shift_date), 101) +' 22:29:59' then A.amnt_OK else 0 end) act_qty_jam_ke_13 ")
        query.AppendLine(" 			,(case when A.time_sagyo between CONVERT(varchar, convert(date,A.shift_date), 101) +' 21:30:00' and CONVERT(varchar, convert(date,A.shift_date), 101) +' 22:29:59' then A.amnt_OK else 0 end) * isnull(D.price, 0) act_amount_jam_ke_13 ")
        query.AppendLine(" 			,(case when A.time_sagyo between CONVERT(varchar, convert(date,A.shift_date), 101) +' 22:30:00' and CONVERT(varchar, convert(date,A.shift_date), 101) +' 23:29:59' then A.amnt_OK else 0 end) act_qty_jam_ke_14 ")
        query.AppendLine(" 			,(case when A.time_sagyo between CONVERT(varchar, convert(date,A.shift_date), 101) +' 22:30:00' and CONVERT(varchar, convert(date,A.shift_date), 101) +' 23:29:59' then A.amnt_OK else 0 end) * isnull(D.price, 0) act_amount_jam_ke_14 ")
        query.AppendLine(" 			,(case when A.time_sagyo between CONVERT(varchar, convert(date,A.shift_date), 101) +' 23:30:00' and CONVERT(varchar, convert(date,A.shift_date), 101) +' 23:59:59' then A.amnt_OK else 0 end) act_qty_jam_ke_15 ")
        query.AppendLine(" 			,(case when A.time_sagyo between CONVERT(varchar, convert(date,A.shift_date), 101) +' 23:30:00' and CONVERT(varchar, convert(date,A.shift_date), 101) +' 23:59:59' then A.amnt_OK else 0 end) * isnull(D.price, 0) act_amount_jam_ke_15 ")
        query.AppendLine(" 			,(case when A.time_sagyo between CONVERT(varchar, DATEADD(day, 1, convert(date,A.shift_date)), 101) +' 00:00:00' and CONVERT(varchar, DATEADD(day, 1, convert(date,A.shift_date)), 101) +' 00:29:59' then A.amnt_OK else 0 end) act_qty_jam_ke_16 ")
        query.AppendLine(" 			,(case when A.time_sagyo between CONVERT(varchar, DATEADD(day, 1, convert(date,A.shift_date)), 101) +' 00:00:00' and CONVERT(varchar, DATEADD(day, 1, convert(date,A.shift_date)), 101) +' 00:29:59' then A.amnt_OK else 0 end) * isnull(D.price, 0) act_amount_jam_ke_16 ")
        query.AppendLine(" 			,(case when A.time_sagyo between CONVERT(varchar, DATEADD(day, 1, convert(date,A.shift_date)), 101) +' 00:30:00' and CONVERT(varchar, DATEADD(day, 1, convert(date,A.shift_date)), 101) +' 01:29:59' then A.amnt_OK else 0 end) act_qty_jam_ke_17 ")
        query.AppendLine(" 			,(case when A.time_sagyo between CONVERT(varchar, DATEADD(day, 1, convert(date,A.shift_date)), 101) +' 00:30:00' and CONVERT(varchar, DATEADD(day, 1, convert(date,A.shift_date)), 101) +' 01:29:59' then A.amnt_OK else 0 end) * isnull(D.price, 0) act_amount_jam_ke_17 ")
        query.AppendLine(" 			,(case when A.time_sagyo between CONVERT(varchar, DATEADD(day, 1, convert(date,A.shift_date)), 101) +' 01:30:00' and CONVERT(varchar, DATEADD(day, 1, convert(date,A.shift_date)), 101) +' 02:29:59' then A.amnt_OK else 0 end) act_qty_jam_ke_18 ")
        query.AppendLine(" 			,(case when A.time_sagyo between CONVERT(varchar, DATEADD(day, 1, convert(date,A.shift_date)), 101) +' 01:30:00' and CONVERT(varchar, DATEADD(day, 1, convert(date,A.shift_date)), 101) +' 02:29:59' then A.amnt_OK else 0 end) * isnull(D.price, 0) act_amount_jam_ke_18 ")
        query.AppendLine(" 			,(case when A.time_sagyo between CONVERT(varchar, DATEADD(day, 1, convert(date,A.shift_date)), 101) +' 02:30:00' and CONVERT(varchar, DATEADD(day, 1, convert(date,A.shift_date)), 101) +' 03:29:59' then A.amnt_OK else 0 end) act_qty_jam_ke_19 ")
        query.AppendLine(" 			,(case when A.time_sagyo between CONVERT(varchar, DATEADD(day, 1, convert(date,A.shift_date)), 101) +' 02:30:00' and CONVERT(varchar, DATEADD(day, 1, convert(date,A.shift_date)), 101) +' 03:29:59' then A.amnt_OK else 0 end) * isnull(D.price, 0) act_amount_jam_ke_19 ")
        query.AppendLine(" 			,(case when A.time_sagyo between CONVERT(varchar, DATEADD(day, 1, convert(date,A.shift_date)), 101) +' 03:30:00' and CONVERT(varchar, DATEADD(day, 1, convert(date,A.shift_date)), 101) +' 05:29:59' then A.amnt_OK else 0 end) act_qty_jam_ke_20 ")
        query.AppendLine(" 			,(case when A.time_sagyo between CONVERT(varchar, DATEADD(day, 1, convert(date,A.shift_date)), 101) +' 03:30:00' and CONVERT(varchar, DATEADD(day, 1, convert(date,A.shift_date)), 101) +' 05:29:59' then A.amnt_OK else 0 end) * isnull(D.price, 0) act_amount_jam_ke_20 ")
        query.AppendLine(" 			,(case when A.time_sagyo between CONVERT(varchar, DATEADD(day, 1, convert(date,A.shift_date)), 101) +' 05:30:00' and CONVERT(varchar, DATEADD(day, 1, convert(date,A.shift_date)), 101) +' 06:29:59' then A.amnt_OK else 0 end) act_qty_jam_ke_21 ")
        query.AppendLine(" 			,(case when A.time_sagyo between CONVERT(varchar, DATEADD(day, 1, convert(date,A.shift_date)), 101) +' 05:30:00' and CONVERT(varchar, DATEADD(day, 1, convert(date,A.shift_date)), 101) +' 06:29:59' then A.amnt_OK else 0 end) * isnull(D.price, 0) act_amount_jam_ke_21 ")
        query.AppendLine(" 			,(case when A.time_sagyo between CONVERT(varchar, DATEADD(day, 1, convert(date,A.shift_date)), 101) +' 06:30:00' and CONVERT(varchar, DATEADD(day, 1, convert(date,A.shift_date)), 101) +' 07:29:59' then A.amnt_OK else 0 end) act_qty_jam_ke_22 ")
        query.AppendLine(" 			,(case when A.time_sagyo between CONVERT(varchar, DATEADD(day, 1, convert(date,A.shift_date)), 101) +' 06:30:00' and CONVERT(varchar, DATEADD(day, 1, convert(date,A.shift_date)), 101) +' 07:29:59' then A.amnt_OK else 0 end) * isnull(D.price, 0) act_amount_jam_ke_22 ")
        query.AppendLine(" 			,A.time_dandori SORT_DATE ")
        query.AppendLine(" 			,isnull(C.id_juchuchuban, '') ORDER_NO ")
        query.AppendLine(" 			,isnull(D.price, 0) PRICE ")
        query.AppendLine(" 			,isnull(D.price, 0) * A.amnt_OK AMOUNT ")
        query.AppendLine(" 			,convert(decimal(8,2),isnull(NULLIF(trim(E.DM_SIZE),''), 0)) * A.amnt_OK TOTAL_SIZE --dibuat nilai menjadi NULL jika hanya space ")
        query.AppendLine(" 		from ")
        query.AppendLine(" 			Z_RT_data_J_kotei A ")
        query.AppendLine(" 			inner join Z_RT_data_J_seisanID B ON A.id_seisan = B.id_seisan ")
        query.AppendLine(" 			inner join Z_RT_data_K_seisanrenban C ON B.id_seisanrenban = C.id_seisanrenban ")
        'menggunakan harga terbaru berdasarkan tanggal pickup plan, menggunakan pickup plan karena jika menggunakan pickup actual nilai defaultnya tahun 2000 jika aktual belum di pickup 
        'query.AppendLine(" 			inner join ( select i_salesorder it_no, max(v_order) price from TPICSDTI.TxDTIPRD.dbo.y_salesorder_addon group by i_salesorder ) as D on D.it_no = C.id_juchuchuban ")
        query.AppendLine(" 			inner join ( ")
        query.AppendLine(" 			             select ")
        query.AppendLine(" 			             	it_no, ")
        query.AppendLine(" 			             	d_pickupplan, ")
        query.AppendLine(" 			             	price, ")
        query.AppendLine(" 			             	RowNum ")
        query.AppendLine(" 			             from ( ")
        query.AppendLine(" 			                    select ")
        query.AppendLine(" 			                        i_salesorder it_no, ")
        query.AppendLine(" 			                        d_pickupplan, ")
        query.AppendLine(" 			                        v_order price, ")
        query.AppendLine(" 			                        row_number() over(partition by i_salesorder order by d_pickupplan desc) as RowNum ")
        query.AppendLine(" 			                    from ")
        query.AppendLine(" 			                        TPICSDTI.TxDTIPRD.dbo.y_salesorder_addon ")
        query.AppendLine(" 			                   ) as A ")
        query.AppendLine(" 			             where RowNum = 1 ")
        query.AppendLine(" 			            ) as D on D.it_no = C.id_juchuchuban ")
        query.AppendLine(" 			left join TPICSDTI.TxDTIPRD.dbo.XITEM E ON A.id_seihin = E.CODE ")
        query.AppendLine(" 		where ")
        query.AppendLine(" 			(B.id_seihin <> 'NOT_APPLY' OR B.id_seihin IS NOT NULL) ")
        'query.AppendLine(" 			--and (A.flg_sagyokanryo = 1) ")
        query.AppendLine(" 			and (A.shift_date between '" + start_date.ToString("yyyyMMdd") + "' and '" + end_date.ToString("yyyyMMdd") + "') ")
        query.AppendLine(" 			and (A.id_kotei = 5230) ")
        query.AppendLine(" 	) as tbl_qty_amount ")
        query.AppendLine(" group by DATE ")
        query.AppendLine(" order by DATE ")

        Dim dt_sum As DataTable
        Try
            'Binding query
            dt_sum = ClsConfig.ExecuteQuery(query.ToString(), ClsConfig.IPServer_RTJN_PRD)

            'Jika periode bulan saat ini belum ada data --> ambil periode bulan sebelumnya.
            'If (dt_sum.Rows.Count = 0) Then
            '    start_date = start_date.AddMonths(-1)
            '    end_date = end_date.AddMonths(-1)
            '    end_date = ClsGeneral.get_last_date(end_date)

            '    get_calculation(start_date, end_date, currentDate, currentHHmm)
            'End If

        Catch ex As Exception
            ClsConfig.create_log_error(currentDate, jenis_laporan, "[" + DateTime.Now.ToString("dd-MM-yyyy hh:mm:ss") + "] -- [ " + ex.Message + " ] -- Get Data Error")
            Environment.Exit(0)
        End Try
        query.Length = 0
        query.Capacity = 0

        Console.WriteLine("##FINISH GET DATA")

        Console.WriteLine("")
        Console.WriteLine("Total Calculation Record : " & dt_sum.Rows.Count)

        Console.WriteLine("")

        Console.WriteLine("##START CALCULATION")
        Try

            If dt_sum.Rows.Count > 0 Then

                Dim day = start_date
                Dim end_day = end_date

                Dim str_date As String = ""
                Dim str_target_print_qty As Int32 = 0

                Dim str_target_amount_day As Double = 0
                Dim str_target_amount_akum As Double = 0
                Dim str_target_qty_day As Int32 = 0
                Dim str_target_qty_akum As Int32 = 0

                Dim str_actual_amount_day As Double = 0
                Dim str_actual_amount_akum As Double = 0
                Dim str_dif_amount_target_actual As Double = 0
                Dim str_actual_qty_day As Int32 = 0
                Dim str_actual_qty_akum As Int32 = 0
                Dim str_dif_qty_target_actual As Int32 = 0

                Dim str_average_size As Double = 0

                Dim act_qty_jam_ke_1 As Int32 = 0
                Dim act_amount_jam_ke_1 As Double = 0
                Dim act_amount_jam_ke_1_shift As Int32
                Dim act_amount_jam_ke_1_group As String

                Dim act_qty_jam_ke_2 As Int32 = 0
                Dim act_amount_jam_ke_2 As Double = 0
                Dim act_amount_jam_ke_2_shift As Int32
                Dim act_amount_jam_ke_2_group As String

                Dim act_qty_jam_ke_3 As Int32 = 0
                Dim act_amount_jam_ke_3 As Double = 0
                Dim act_amount_jam_ke_3_shift As Int32
                Dim act_amount_jam_ke_3_group As String

                Dim act_qty_jam_ke_4 As Int32 = 0
                Dim act_amount_jam_ke_4 As Double = 0
                Dim act_amount_jam_ke_4_shift As Int32
                Dim act_amount_jam_ke_4_group As String

                Dim act_qty_jam_ke_5 As Int32 = 0
                Dim act_amount_jam_ke_5 As Double = 0
                Dim act_amount_jam_ke_5_shift As Int32
                Dim act_amount_jam_ke_5_group As String

                Dim act_qty_jam_ke_6 As Int32 = 0
                Dim act_amount_jam_ke_6 As Double = 0
                Dim act_amount_jam_ke_6_shift As Int32
                Dim act_amount_jam_ke_6_group As String

                Dim act_qty_jam_ke_7 As Int32 = 0
                Dim act_amount_jam_ke_7 As Double = 0
                Dim act_amount_jam_ke_7_shift As Int32
                Dim act_amount_jam_ke_7_group As String

                Dim act_qty_jam_ke_8 As Int32 = 0
                Dim act_amount_jam_ke_8 As Double = 0
                Dim act_amount_jam_ke_8_shift As Int32
                Dim act_amount_jam_ke_8_group As String

                Dim act_qty_jam_ke_9 As Int32 = 0
                Dim act_amount_jam_ke_9 As Double = 0
                Dim act_amount_jam_ke_9_shift As Int32
                Dim act_amount_jam_ke_9_group As String

                Dim act_qty_jam_ke_10 As Int32 = 0
                Dim act_amount_jam_ke_10 As Double = 0
                Dim act_amount_jam_ke_10_shift As Int32
                Dim act_amount_jam_ke_10_group As String

                Dim act_qty_jam_ke_11 As Int32 = 0
                Dim act_amount_jam_ke_11 As Double = 0
                Dim act_amount_jam_ke_11_shift As Int32
                Dim act_amount_jam_ke_11_group As String

                Dim act_qty_jam_ke_12 As Int32 = 0
                Dim act_amount_jam_ke_12 As Double = 0
                Dim act_amount_jam_ke_12_shift As Int32
                Dim act_amount_jam_ke_12_group As String

                Dim act_qty_jam_ke_13 As Int32 = 0
                Dim act_amount_jam_ke_13 As Double = 0
                Dim act_amount_jam_ke_13_shift As Int32
                Dim act_amount_jam_ke_13_group As String

                Dim act_qty_jam_ke_14 As Int32 = 0
                Dim act_amount_jam_ke_14 As Double = 0
                Dim act_amount_jam_ke_14_shift As Int32
                Dim act_amount_jam_ke_14_group As String

                Dim act_qty_jam_ke_15_16_istirahat As Int32 = 0
                Dim act_amount_jam_ke_15_16_istirahat As Double = 0
                Dim act_amount_jam_ke_15_16_istirahat_shift As Int32
                Dim act_amount_jam_ke_15_16_istirahat_group As String

                Dim act_qty_jam_ke_17 As Int32 = 0
                Dim act_amount_jam_ke_17 As Double = 0
                Dim act_amount_jam_ke_17_shift As Int32
                Dim act_amount_jam_ke_17_group As String

                Dim act_qty_jam_ke_18 As Int32 = 0
                Dim act_amount_jam_ke_18 As Double = 0
                Dim act_amount_jam_ke_18_shift As Int32
                Dim act_amount_jam_ke_18_group As String

                Dim act_qty_jam_ke_19 As Int32 = 0
                Dim act_amount_jam_ke_19 As Double = 0
                Dim act_amount_jam_ke_19_shift As Int32
                Dim act_amount_jam_ke_19_group As String

                Dim act_qty_jam_ke_20 As Int32 = 0
                Dim act_amount_jam_ke_20 As Double = 0
                Dim act_amount_jam_ke_20_shift As Int32
                Dim act_amount_jam_ke_20_group As String

                Dim act_qty_jam_ke_21 As Int32 = 0
                Dim act_amount_jam_ke_21 As Double = 0
                Dim act_amount_jam_ke_21_shift As Int32
                Dim act_amount_jam_ke_21_group As String

                Dim act_qty_jam_ke_22 As Int32 = 0
                Dim act_amount_jam_ke_22 As Double = 0
                Dim act_amount_jam_ke_22_shift As Int32
                Dim act_amount_jam_ke_22_group As String

                While day <= end_day
                    Console.WriteLine("Calculation Date : " & day.ToString("dd-MM-yyyy"))

                    For i = 0 To dt_sum.Rows.Count - 1
                        If day.ToString("yyyy-MM-dd") = DateTime.Parse(dt_sum(i)("DATE")).ToString("yyyy-MM-dd") Then

                            str_date = day.ToString("yyyy-MM-dd")

                            act_qty_jam_ke_1 = Int32.Parse(dt_sum(i)("act_qty_jam_ke_1").ToString())
                            act_amount_jam_ke_1 = Double.Parse(dt_sum(i)("act_amount_jam_ke_1").ToString()) / 1000000
                            act_amount_jam_ke_1_shift = 1
                            act_amount_jam_ke_1_group = ClsGeneral.find_shift_group(day.ToString("yyyy-MM-dd"), act_amount_jam_ke_1_shift)

                            act_qty_jam_ke_2 = Int32.Parse(dt_sum(i)("act_qty_jam_ke_2").ToString())
                            act_amount_jam_ke_2 = Double.Parse(dt_sum(i)("act_amount_jam_ke_2").ToString()) / 1000000
                            act_amount_jam_ke_2_shift = 1
                            act_amount_jam_ke_2_group = ClsGeneral.find_shift_group(day.ToString("yyyy-MM-dd"), act_amount_jam_ke_2_shift)

                            act_qty_jam_ke_3 = Int32.Parse(dt_sum(i)("act_qty_jam_ke_3").ToString())
                            act_amount_jam_ke_3 = Double.Parse(dt_sum(i)("act_amount_jam_ke_3").ToString()) / 1000000
                            act_amount_jam_ke_3_shift = 1
                            act_amount_jam_ke_3_group = ClsGeneral.find_shift_group(day.ToString("yyyy-MM-dd"), act_amount_jam_ke_3_shift)

                            act_qty_jam_ke_4 = Int32.Parse(dt_sum(i)("act_qty_jam_ke_4").ToString())
                            act_amount_jam_ke_4 = Double.Parse(dt_sum(i)("act_amount_jam_ke_4").ToString()) / 1000000
                            act_amount_jam_ke_4_shift = 1
                            act_amount_jam_ke_4_group = ClsGeneral.find_shift_group(day.ToString("yyyy-MM-dd"), act_amount_jam_ke_4_shift)

                            act_qty_jam_ke_5 = Int32.Parse(dt_sum(i)("act_qty_jam_ke_5").ToString())
                            act_amount_jam_ke_5 = Double.Parse(dt_sum(i)("act_amount_jam_ke_5").ToString()) / 1000000
                            act_amount_jam_ke_5_shift = 1
                            act_amount_jam_ke_5_group = ClsGeneral.find_shift_group(day.ToString("yyyy-MM-dd"), act_amount_jam_ke_5_shift)

                            act_qty_jam_ke_6 = Int32.Parse(dt_sum(i)("act_qty_jam_ke_6").ToString())
                            act_amount_jam_ke_6 = Double.Parse(dt_sum(i)("act_amount_jam_ke_6").ToString()) / 1000000
                            act_amount_jam_ke_6_shift = 1
                            act_amount_jam_ke_6_group = ClsGeneral.find_shift_group(day.ToString("yyyy-MM-dd"), act_amount_jam_ke_6_shift)

                            act_qty_jam_ke_7 = Int32.Parse(dt_sum(i)("act_qty_jam_ke_7").ToString())
                            act_amount_jam_ke_7 = Double.Parse(dt_sum(i)("act_amount_jam_ke_7").ToString()) / 1000000
                            act_amount_jam_ke_7_shift = 1
                            act_amount_jam_ke_7_group = ClsGeneral.find_shift_group(day.ToString("yyyy-MM-dd"), act_amount_jam_ke_7_shift)

                            act_qty_jam_ke_8 = Int32.Parse(dt_sum(i)("act_qty_jam_ke_8").ToString())
                            act_amount_jam_ke_8 = Double.Parse(dt_sum(i)("act_amount_jam_ke_8").ToString()) / 1000000
                            act_amount_jam_ke_8_shift = 1
                            act_amount_jam_ke_8_group = ClsGeneral.find_shift_group(day.ToString("yyyy-MM-dd"), act_amount_jam_ke_8_shift)

                            act_qty_jam_ke_9 = Int32.Parse(dt_sum(i)("act_qty_jam_ke_9").ToString())
                            act_amount_jam_ke_9 = Double.Parse(dt_sum(i)("act_amount_jam_ke_9").ToString()) / 1000000
                            act_amount_jam_ke_9_shift = 1
                            act_amount_jam_ke_9_group = ClsGeneral.find_shift_group(day.ToString("yyyy-MM-dd"), act_amount_jam_ke_9_shift)

                            act_qty_jam_ke_10 = Int32.Parse(dt_sum(i)("act_qty_jam_ke_10").ToString())
                            act_amount_jam_ke_10 = Double.Parse(dt_sum(i)("act_amount_jam_ke_10").ToString()) / 1000000
                            act_amount_jam_ke_10_shift = 1
                            act_amount_jam_ke_10_group = ClsGeneral.find_shift_group(day.ToString("yyyy-MM-dd"), act_amount_jam_ke_10_shift)

                            act_qty_jam_ke_11 = Int32.Parse(dt_sum(i)("act_qty_jam_ke_11").ToString())
                            act_amount_jam_ke_11 = Double.Parse(dt_sum(i)("act_amount_jam_ke_11").ToString()) / 1000000
                            act_amount_jam_ke_11_shift = 2
                            act_amount_jam_ke_11_group = ClsGeneral.find_shift_group(day.ToString("yyyy-MM-dd"), act_amount_jam_ke_11_shift)

                            act_qty_jam_ke_12 = Int32.Parse(dt_sum(i)("act_qty_jam_ke_12").ToString())
                            act_amount_jam_ke_12 = Double.Parse(dt_sum(i)("act_amount_jam_ke_12").ToString()) / 1000000
                            act_amount_jam_ke_12_shift = 2
                            act_amount_jam_ke_12_group = ClsGeneral.find_shift_group(day.ToString("yyyy-MM-dd"), act_amount_jam_ke_12_shift)

                            act_qty_jam_ke_13 = Int32.Parse(dt_sum(i)("act_qty_jam_ke_13").ToString())
                            act_amount_jam_ke_13 = Double.Parse(dt_sum(i)("act_amount_jam_ke_13").ToString()) / 1000000
                            act_amount_jam_ke_13_shift = 2
                            act_amount_jam_ke_13_group = ClsGeneral.find_shift_group(day.ToString("yyyy-MM-dd"), act_amount_jam_ke_13_shift)

                            act_qty_jam_ke_14 = Int32.Parse(dt_sum(i)("act_qty_jam_ke_14").ToString())
                            act_amount_jam_ke_14 = Double.Parse(dt_sum(i)("act_amount_jam_ke_14").ToString()) / 1000000
                            act_amount_jam_ke_14_shift = 2
                            act_amount_jam_ke_14_group = ClsGeneral.find_shift_group(day.ToString("yyyy-MM-dd"), act_amount_jam_ke_14_shift)

                            act_qty_jam_ke_15_16_istirahat = Int32.Parse(dt_sum(i)("act_qty_jam_ke_15").ToString()) + Int32.Parse(dt_sum(i)("act_qty_jam_ke_16").ToString())
                            act_amount_jam_ke_15_16_istirahat = (Double.Parse(dt_sum(i)("act_amount_jam_ke_15").ToString()) + Double.Parse(dt_sum(i)("act_amount_jam_ke_16").ToString())) / 1000000
                            act_amount_jam_ke_15_16_istirahat_shift = 2
                            act_amount_jam_ke_15_16_istirahat_group = ClsGeneral.find_shift_group(day.ToString("yyyy-MM-dd"), act_amount_jam_ke_15_16_istirahat_shift)

                            act_qty_jam_ke_17 = Int32.Parse(dt_sum(i)("act_qty_jam_ke_17").ToString())
                            act_amount_jam_ke_17 = Double.Parse(dt_sum(i)("act_amount_jam_ke_17").ToString()) / 1000000
                            act_amount_jam_ke_17_shift = 2
                            act_amount_jam_ke_17_group = ClsGeneral.find_shift_group(day.ToString("yyyy-MM-dd"), act_amount_jam_ke_17_shift)

                            act_qty_jam_ke_18 = Int32.Parse(dt_sum(i)("act_qty_jam_ke_18").ToString())
                            act_amount_jam_ke_18 = Double.Parse(dt_sum(i)("act_amount_jam_ke_18").ToString()) / 1000000
                            act_amount_jam_ke_18_shift = 2
                            act_amount_jam_ke_18_group = ClsGeneral.find_shift_group(day.ToString("yyyy-MM-dd"), act_amount_jam_ke_18_shift)

                            act_qty_jam_ke_19 = Int32.Parse(dt_sum(i)("act_qty_jam_ke_19").ToString())
                            act_amount_jam_ke_19 = Double.Parse(dt_sum(i)("act_amount_jam_ke_19").ToString()) / 1000000
                            act_amount_jam_ke_19_shift = 2
                            act_amount_jam_ke_19_group = ClsGeneral.find_shift_group(day.ToString("yyyy-MM-dd"), act_amount_jam_ke_19_shift)

                            act_qty_jam_ke_20 = Int32.Parse(dt_sum(i)("act_qty_jam_ke_20").ToString())
                            act_amount_jam_ke_20 = Double.Parse(dt_sum(i)("act_amount_jam_ke_20").ToString()) / 1000000
                            act_amount_jam_ke_20_shift = 2
                            act_amount_jam_ke_20_group = ClsGeneral.find_shift_group(day.ToString("yyyy-MM-dd"), act_amount_jam_ke_20_shift)

                            act_qty_jam_ke_21 = Int32.Parse(dt_sum(i)("act_qty_jam_ke_21").ToString())
                            act_amount_jam_ke_21 = Double.Parse(dt_sum(i)("act_amount_jam_ke_21").ToString()) / 1000000
                            act_amount_jam_ke_21_shift = 2
                            act_amount_jam_ke_21_group = ClsGeneral.find_shift_group(day.ToString("yyyy-MM-dd"), act_amount_jam_ke_21_shift)

                            act_qty_jam_ke_22 = Int32.Parse(dt_sum(i)("act_qty_jam_ke_22").ToString())
                            act_amount_jam_ke_22 = Double.Parse(dt_sum(i)("act_amount_jam_ke_22").ToString()) / 1000000
                            act_amount_jam_ke_22_shift = 2
                            act_amount_jam_ke_22_group = ClsGeneral.find_shift_group(day.ToString("yyyy-MM-dd"), act_amount_jam_ke_22_shift)

                            str_target_print_qty = ClsGeneral.find_target_print_qty(day.ToString("yyyy-MM-dd"))

                            str_target_qty_day = ClsGeneral.find_target_qty(day.ToString("yyyy-MM-dd"))
                            str_target_qty_akum = str_target_qty_akum + str_target_qty_day

                            str_target_amount_day = ClsGeneral.find_target_amount(day.ToString("yyyy-MM-dd"))
                            str_target_amount_akum = str_target_amount_akum + str_target_amount_day

                            str_actual_qty_day = (
                                                        act_qty_jam_ke_1 + act_qty_jam_ke_2 + act_qty_jam_ke_3 _
                                                        + act_qty_jam_ke_4 + act_qty_jam_ke_5 + act_qty_jam_ke_6 _
                                                        + act_qty_jam_ke_7 + act_qty_jam_ke_8 + act_qty_jam_ke_9 _
                                                        + act_qty_jam_ke_10 + act_qty_jam_ke_11 + act_qty_jam_ke_12 _
                                                        + act_qty_jam_ke_13 + act_qty_jam_ke_14 + act_qty_jam_ke_15_16_istirahat _
                                                        + act_qty_jam_ke_17 + act_qty_jam_ke_18 + act_qty_jam_ke_19 _
                                                        + act_qty_jam_ke_20 + act_qty_jam_ke_21 + act_qty_jam_ke_22
                                                    )
                            str_actual_qty_akum = str_actual_qty_akum + str_actual_qty_day
                            str_dif_qty_target_actual = str_actual_qty_akum - str_target_qty_akum

                            str_actual_amount_day = (
                                                        act_amount_jam_ke_1 + act_amount_jam_ke_2 + act_amount_jam_ke_3 _
                                                        + act_amount_jam_ke_4 + act_amount_jam_ke_5 + act_amount_jam_ke_6 _
                                                        + act_amount_jam_ke_7 + act_amount_jam_ke_8 + act_amount_jam_ke_9 _
                                                        + act_amount_jam_ke_10 + act_amount_jam_ke_11 + act_amount_jam_ke_12 _
                                                        + act_amount_jam_ke_13 + act_amount_jam_ke_14 + act_amount_jam_ke_15_16_istirahat _
                                                        + act_amount_jam_ke_17 + act_amount_jam_ke_18 + act_amount_jam_ke_19 _
                                                        + act_amount_jam_ke_20 + act_amount_jam_ke_21 + act_amount_jam_ke_22
                                                    )
                            str_actual_amount_akum = str_actual_amount_akum + str_actual_amount_day
                            str_dif_amount_target_actual = str_actual_amount_akum - str_target_amount_akum

                            If (Double.Parse(dt_sum(i)("TOTAL_SIZE").ToString()) <> "0" And str_actual_qty_day <> 0) Then
                                str_average_size = Double.Parse(dt_sum(i)("TOTAL_SIZE").ToString()) / str_actual_qty_day
                            Else
                                str_average_size = 0
                            End If

                            query.AppendLine(" select ")
                            query.AppendLine("     id ")
                            query.AppendLine("     ,date ")
                            query.AppendLine("     ,target_print_qty")
                            query.AppendLine("     ,target_amount_day ")
                            query.AppendLine("     ,target_amount_akum ")
                            query.AppendLine("     ,actual_amount_day ")
                            query.AppendLine("     ,actual_amount_akum ")
                            query.AppendLine("     ,dif_amount_target_actual ")
                            query.AppendLine("     ,target_qty_day ")
                            query.AppendLine("     ,target_qty_akum ")
                            query.AppendLine("     ,actual_qty_day ")
                            query.AppendLine("     ,actual_qty_akum ")
                            query.AppendLine("     ,dif_qty_target_actual ")
                            query.AppendLine("     ,average_size ")
                            query.AppendLine(" from ")
                            query.AppendLine("     ad_dis_rtjn_sum_qty_amount ")
                            query.AppendLine(" where ")
                            query.AppendLine("     date = '" & str_date & "' ")
                            dt = ClsConfig.ExecuteQuery(query.ToString(), ClsConfig.IPServer_ADDONS)
                            query.Length = 0
                            query.Capacity = 0

                            If dt.Rows.Count > 0 Then
                                'If day.ToString("yyyy-MM-dd") = Today.ToString("yyyy-MM-dd") Or day.ToString("yyyy-MM-dd") = Today.AddDays(-1).ToString("yyyy-MM-dd") Then 'update hanya 2 hari (hari ini dan kemarin)                            
                                query.AppendLine(" update ")
                                query.AppendLine("     ad_dis_rtjn_sum_qty_amount ")
                                query.AppendLine(" set ")
                                query.AppendLine("     target_print_qty='" & str_target_print_qty.ToString() & "' ")

                                query.AppendLine("     ,target_amount_day='" & str_target_amount_day.ToString() & "' ")
                                query.AppendLine("     ,target_amount_akum='" & str_target_amount_akum.ToString() & "' ")
                                query.AppendLine("     ,actual_amount_day='" & str_actual_amount_day.ToString() & "' ")
                                query.AppendLine("     ,actual_amount_akum='" & str_actual_amount_akum.ToString() & "' ")
                                query.AppendLine("     ,dif_amount_target_actual='" & str_dif_amount_target_actual.ToString() & "' ")

                                query.AppendLine("     ,target_qty_day='" & str_target_qty_day.ToString() & "' ")
                                query.AppendLine("     ,target_qty_akum='" & str_target_qty_akum.ToString() & "' ")
                                query.AppendLine("     ,actual_qty_day='" & str_actual_qty_day.ToString() & "' ")
                                query.AppendLine("     ,actual_qty_akum='" & str_actual_qty_akum.ToString() & "' ")
                                query.AppendLine("     ,dif_qty_target_actual='" & str_dif_qty_target_actual.ToString() & "' ")

                                query.AppendLine("     ,average_size='" & str_average_size.ToString() & "' ")

                                query.AppendLine("     ,act_qty_jam_ke_1='" & act_qty_jam_ke_1.ToString() & "' ")
                                query.AppendLine("     ,act_amount_jam_ke_1='" & act_amount_jam_ke_1.ToString() & "' ")
                                query.AppendLine("     ,act_amount_jam_ke_1_shift='" & act_amount_jam_ke_1_shift.ToString() & "' ")
                                query.AppendLine("     ,act_amount_jam_ke_1_group='" & act_amount_jam_ke_1_group.ToString() & "' ")

                                query.AppendLine("     ,act_qty_jam_ke_2='" & act_qty_jam_ke_2.ToString() & "' ")
                                query.AppendLine("     ,act_amount_jam_ke_2='" & act_amount_jam_ke_2.ToString() & "' ")
                                query.AppendLine("     ,act_amount_jam_ke_2_shift='" & act_amount_jam_ke_2_shift.ToString() & "' ")
                                query.AppendLine("     ,act_amount_jam_ke_2_group='" & act_amount_jam_ke_2_group.ToString() & "' ")

                                query.AppendLine("     ,act_qty_jam_ke_3='" & act_qty_jam_ke_3.ToString() & "' ")
                                query.AppendLine("     ,act_amount_jam_ke_3='" & act_amount_jam_ke_3.ToString() & "' ")
                                query.AppendLine("     ,act_amount_jam_ke_3_shift='" & act_amount_jam_ke_3_shift.ToString() & "' ")
                                query.AppendLine("     ,act_amount_jam_ke_3_group='" & act_amount_jam_ke_3_group.ToString() & "' ")

                                query.AppendLine("     ,act_qty_jam_ke_4='" & act_qty_jam_ke_4.ToString() & "' ")
                                query.AppendLine("     ,act_amount_jam_ke_4='" & act_amount_jam_ke_4.ToString() & "' ")
                                query.AppendLine("     ,act_amount_jam_ke_4_shift='" & act_amount_jam_ke_4_shift.ToString() & "' ")
                                query.AppendLine("     ,act_amount_jam_ke_4_group='" & act_amount_jam_ke_4_group.ToString() & "' ")

                                query.AppendLine("     ,act_qty_jam_ke_5='" & act_qty_jam_ke_5.ToString() & "' ")
                                query.AppendLine("     ,act_amount_jam_ke_5='" & act_amount_jam_ke_5.ToString() & "' ")
                                query.AppendLine("     ,act_amount_jam_ke_5_shift='" & act_amount_jam_ke_5_shift.ToString() & "' ")
                                query.AppendLine("     ,act_amount_jam_ke_5_group='" & act_amount_jam_ke_5_group.ToString() & "' ")

                                query.AppendLine("     ,act_qty_jam_ke_6='" & act_qty_jam_ke_6.ToString() & "' ")
                                query.AppendLine("     ,act_amount_jam_ke_6='" & act_amount_jam_ke_6.ToString() & "' ")
                                query.AppendLine("     ,act_amount_jam_ke_6_shift='" & act_amount_jam_ke_6_shift.ToString() & "' ")
                                query.AppendLine("     ,act_amount_jam_ke_6_group='" & act_amount_jam_ke_6_group.ToString() & "' ")

                                query.AppendLine("     ,act_qty_jam_ke_7='" & act_qty_jam_ke_7.ToString() & "' ")
                                query.AppendLine("     ,act_amount_jam_ke_7='" & act_amount_jam_ke_7.ToString() & "' ")
                                query.AppendLine("     ,act_amount_jam_ke_7_shift='" & act_amount_jam_ke_7_shift.ToString() & "' ")
                                query.AppendLine("     ,act_amount_jam_ke_7_group='" & act_amount_jam_ke_7_group.ToString() & "' ")

                                query.AppendLine("     ,act_qty_jam_ke_8='" & act_qty_jam_ke_8.ToString() & "' ")
                                query.AppendLine("     ,act_amount_jam_ke_8='" & act_amount_jam_ke_8.ToString() & "' ")
                                query.AppendLine("     ,act_amount_jam_ke_8_shift='" & act_amount_jam_ke_8_shift.ToString() & "' ")
                                query.AppendLine("     ,act_amount_jam_ke_8_group='" & act_amount_jam_ke_8_group.ToString() & "' ")

                                query.AppendLine("     ,act_qty_jam_ke_9='" & act_qty_jam_ke_9.ToString() & "' ")
                                query.AppendLine("     ,act_amount_jam_ke_9='" & act_amount_jam_ke_9.ToString() & "' ")
                                query.AppendLine("     ,act_amount_jam_ke_9_shift='" & act_amount_jam_ke_9_shift.ToString() & "' ")
                                query.AppendLine("     ,act_amount_jam_ke_9_group='" & act_amount_jam_ke_9_group.ToString() & "' ")

                                query.AppendLine("     ,act_qty_jam_ke_10='" & act_qty_jam_ke_10.ToString() & "' ")
                                query.AppendLine("     ,act_amount_jam_ke_10='" & act_amount_jam_ke_10.ToString() & "' ")
                                query.AppendLine("     ,act_amount_jam_ke_10_shift='" & act_amount_jam_ke_10_shift.ToString() & "' ")
                                query.AppendLine("     ,act_amount_jam_ke_10_group='" & act_amount_jam_ke_10_group.ToString() & "' ")

                                query.AppendLine("     ,act_qty_jam_ke_11='" & act_qty_jam_ke_11.ToString() & "' ")
                                query.AppendLine("     ,act_amount_jam_ke_11='" & act_amount_jam_ke_11.ToString() & "' ")
                                query.AppendLine("     ,act_amount_jam_ke_11_shift='" & act_amount_jam_ke_11_shift.ToString() & "' ")
                                query.AppendLine("     ,act_amount_jam_ke_11_group='" & act_amount_jam_ke_11_group.ToString() & "' ")

                                query.AppendLine("     ,act_qty_jam_ke_12='" & act_qty_jam_ke_12.ToString() & "' ")
                                query.AppendLine("     ,act_amount_jam_ke_12='" & act_amount_jam_ke_12.ToString() & "' ")
                                query.AppendLine("     ,act_amount_jam_ke_12_shift='" & act_amount_jam_ke_12_shift.ToString() & "' ")
                                query.AppendLine("     ,act_amount_jam_ke_12_group='" & act_amount_jam_ke_12_group.ToString() & "' ")

                                query.AppendLine("     ,act_qty_jam_ke_13='" & act_qty_jam_ke_13.ToString() & "' ")
                                query.AppendLine("     ,act_amount_jam_ke_13='" & act_amount_jam_ke_13.ToString() & "' ")
                                query.AppendLine("     ,act_amount_jam_ke_13_shift='" & act_amount_jam_ke_13_shift.ToString() & "' ")
                                query.AppendLine("     ,act_amount_jam_ke_13_group='" & act_amount_jam_ke_13_group.ToString() & "' ")

                                query.AppendLine("     ,act_qty_jam_ke_14='" & act_qty_jam_ke_14.ToString() & "' ")
                                query.AppendLine("     ,act_amount_jam_ke_14='" & act_amount_jam_ke_14.ToString() & "' ")
                                query.AppendLine("     ,act_amount_jam_ke_14_shift='" & act_amount_jam_ke_14_shift.ToString() & "' ")
                                query.AppendLine("     ,act_amount_jam_ke_14_group='" & act_amount_jam_ke_14_group.ToString() & "' ")

                                query.AppendLine("     ,act_qty_jam_ke_15_16_istirahat='" & act_qty_jam_ke_15_16_istirahat.ToString() & "' ")
                                query.AppendLine("     ,act_amount_jam_ke_15_16_istirahat='" & act_amount_jam_ke_15_16_istirahat.ToString() & "' ")
                                query.AppendLine("     ,act_amount_jam_ke_15_16_istirahat_shift='" & act_amount_jam_ke_15_16_istirahat_shift.ToString() & "' ")
                                query.AppendLine("     ,act_amount_jam_ke_15_16_istirahat_group='" & act_amount_jam_ke_15_16_istirahat_group.ToString() & "' ")

                                query.AppendLine("     ,act_qty_jam_ke_17='" & act_qty_jam_ke_17.ToString() & "' ")
                                query.AppendLine("     ,act_amount_jam_ke_17='" & act_amount_jam_ke_17.ToString() & "' ")
                                query.AppendLine("     ,act_amount_jam_ke_17_shift='" & act_amount_jam_ke_17_shift.ToString() & "' ")
                                query.AppendLine("     ,act_amount_jam_ke_17_group='" & act_amount_jam_ke_17_group.ToString() & "' ")

                                query.AppendLine("     ,act_qty_jam_ke_18='" & act_qty_jam_ke_18.ToString() & "' ")
                                query.AppendLine("     ,act_amount_jam_ke_18='" & act_amount_jam_ke_18.ToString() & "' ")
                                query.AppendLine("     ,act_amount_jam_ke_18_shift='" & act_amount_jam_ke_18_shift.ToString() & "' ")
                                query.AppendLine("     ,act_amount_jam_ke_18_group='" & act_amount_jam_ke_18_group.ToString() & "' ")

                                query.AppendLine("     ,act_qty_jam_ke_19='" & act_qty_jam_ke_19.ToString() & "' ")
                                query.AppendLine("     ,act_amount_jam_ke_19='" & act_amount_jam_ke_19.ToString() & "' ")
                                query.AppendLine("     ,act_amount_jam_ke_19_shift='" & act_amount_jam_ke_19_shift.ToString() & "' ")
                                query.AppendLine("     ,act_amount_jam_ke_19_group='" & act_amount_jam_ke_19_group.ToString() & "' ")

                                query.AppendLine("     ,act_qty_jam_ke_20='" & act_qty_jam_ke_20.ToString() & "' ")
                                query.AppendLine("     ,act_amount_jam_ke_20='" & act_amount_jam_ke_20.ToString() & "' ")
                                query.AppendLine("     ,act_amount_jam_ke_20_shift='" & act_amount_jam_ke_20_shift.ToString() & "' ")
                                query.AppendLine("     ,act_amount_jam_ke_20_group='" & act_amount_jam_ke_20_group.ToString() & "' ")

                                query.AppendLine("     ,act_qty_jam_ke_21='" & act_qty_jam_ke_21.ToString() & "' ")
                                query.AppendLine("     ,act_amount_jam_ke_21='" & act_amount_jam_ke_21.ToString() & "' ")
                                query.AppendLine("     ,act_amount_jam_ke_21_shift='" & act_amount_jam_ke_21_shift.ToString() & "' ")
                                query.AppendLine("     ,act_amount_jam_ke_21_group='" & act_amount_jam_ke_21_group.ToString() & "' ")

                                query.AppendLine("     ,act_qty_jam_ke_22='" & act_qty_jam_ke_22.ToString() & "' ")
                                query.AppendLine("     ,act_amount_jam_ke_22='" & act_amount_jam_ke_22.ToString() & "' ")
                                query.AppendLine("     ,act_amount_jam_ke_22_shift='" & act_amount_jam_ke_22_shift.ToString() & "' ")
                                query.AppendLine("     ,act_amount_jam_ke_22_group='" & act_amount_jam_ke_22_group.ToString() & "' ")

                                query.AppendLine(" where ")
                                query.AppendLine("     id='" & dt(0)("id").ToString() & "' ")
                                ClsConfig.ExecuteNonQuery(query.ToString(), ClsConfig.IPServer_ADDONS)
                                query.Length = 0
                                query.Capacity = 0
                                Console.WriteLine("Proses Update Data")
                                'End If
                            Else
                                query.AppendLine(" Insert Into ")
                                query.AppendLine("     ad_dis_rtjn_sum_qty_amount ")
                                query.AppendLine("     ( ")
                                query.AppendLine("         date ")
                                query.AppendLine("         ,target_print_qty ")

                                query.AppendLine("         ,target_amount_day ")
                                query.AppendLine("         ,target_amount_akum ")
                                query.AppendLine("         ,actual_amount_day ")
                                query.AppendLine("         ,actual_amount_akum ")
                                query.AppendLine("         ,dif_amount_target_actual ")

                                query.AppendLine("         ,target_qty_day ")
                                query.AppendLine("         ,target_qty_akum ")
                                query.AppendLine("         ,actual_qty_day ")
                                query.AppendLine("         ,actual_qty_akum ")
                                query.AppendLine("         ,dif_qty_target_actual ")

                                query.AppendLine("         ,average_size ")

                                query.AppendLine("         ,act_qty_jam_ke_1 ")
                                query.AppendLine("         ,act_amount_jam_ke_1 ")
                                query.AppendLine("         ,act_amount_jam_ke_1_shift ")
                                query.AppendLine("         ,act_amount_jam_ke_1_group ")

                                query.AppendLine("         ,act_qty_jam_ke_2 ")
                                query.AppendLine("         ,act_amount_jam_ke_2 ")
                                query.AppendLine("         ,act_amount_jam_ke_2_shift ")
                                query.AppendLine("         ,act_amount_jam_ke_2_group ")

                                query.AppendLine("         ,act_qty_jam_ke_3 ")
                                query.AppendLine("         ,act_amount_jam_ke_3 ")
                                query.AppendLine("         ,act_amount_jam_ke_3_shift ")
                                query.AppendLine("         ,act_amount_jam_ke_3_group ")

                                query.AppendLine("         ,act_qty_jam_ke_4 ")
                                query.AppendLine("         ,act_amount_jam_ke_4 ")
                                query.AppendLine("         ,act_amount_jam_ke_4_shift ")
                                query.AppendLine("         ,act_amount_jam_ke_4_group ")

                                query.AppendLine("         ,act_qty_jam_ke_5 ")
                                query.AppendLine("         ,act_amount_jam_ke_5 ")
                                query.AppendLine("         ,act_amount_jam_ke_5_shift ")
                                query.AppendLine("         ,act_amount_jam_ke_5_group ")

                                query.AppendLine("         ,act_qty_jam_ke_6 ")
                                query.AppendLine("         ,act_amount_jam_ke_6 ")
                                query.AppendLine("         ,act_amount_jam_ke_6_shift ")
                                query.AppendLine("         ,act_amount_jam_ke_6_group ")

                                query.AppendLine("         ,act_qty_jam_ke_7 ")
                                query.AppendLine("         ,act_amount_jam_ke_7 ")
                                query.AppendLine("         ,act_amount_jam_ke_7_shift ")
                                query.AppendLine("         ,act_amount_jam_ke_7_group ")

                                query.AppendLine("         ,act_qty_jam_ke_8 ")
                                query.AppendLine("         ,act_amount_jam_ke_8 ")
                                query.AppendLine("         ,act_amount_jam_ke_8_shift ")
                                query.AppendLine("         ,act_amount_jam_ke_8_group ")

                                query.AppendLine("         ,act_qty_jam_ke_9 ")
                                query.AppendLine("         ,act_amount_jam_ke_9 ")
                                query.AppendLine("         ,act_amount_jam_ke_9_shift ")
                                query.AppendLine("         ,act_amount_jam_ke_9_group ")

                                query.AppendLine("         ,act_qty_jam_ke_10 ")
                                query.AppendLine("         ,act_amount_jam_ke_10 ")
                                query.AppendLine("         ,act_amount_jam_ke_10_shift ")
                                query.AppendLine("         ,act_amount_jam_ke_10_group ")

                                query.AppendLine("         ,act_qty_jam_ke_11 ")
                                query.AppendLine("         ,act_amount_jam_ke_11 ")
                                query.AppendLine("         ,act_amount_jam_ke_11_shift ")
                                query.AppendLine("         ,act_amount_jam_ke_11_group ")

                                query.AppendLine("         ,act_qty_jam_ke_12 ")
                                query.AppendLine("         ,act_amount_jam_ke_12 ")
                                query.AppendLine("         ,act_amount_jam_ke_12_shift ")
                                query.AppendLine("         ,act_amount_jam_ke_12_group ")

                                query.AppendLine("         ,act_qty_jam_ke_13 ")
                                query.AppendLine("         ,act_amount_jam_ke_13 ")
                                query.AppendLine("         ,act_amount_jam_ke_13_shift ")
                                query.AppendLine("         ,act_amount_jam_ke_13_group ")

                                query.AppendLine("         ,act_qty_jam_ke_14 ")
                                query.AppendLine("         ,act_amount_jam_ke_14 ")
                                query.AppendLine("         ,act_amount_jam_ke_14_shift ")
                                query.AppendLine("         ,act_amount_jam_ke_14_group ")

                                query.AppendLine("         ,act_qty_jam_ke_15_16_istirahat ")
                                query.AppendLine("         ,act_amount_jam_ke_15_16_istirahat ")
                                query.AppendLine("         ,act_amount_jam_ke_15_16_istirahat_shift ")
                                query.AppendLine("         ,act_amount_jam_ke_15_16_istirahat_group ")

                                query.AppendLine("         ,act_qty_jam_ke_17 ")
                                query.AppendLine("         ,act_amount_jam_ke_17 ")
                                query.AppendLine("         ,act_amount_jam_ke_17_shift ")
                                query.AppendLine("         ,act_amount_jam_ke_17_group ")

                                query.AppendLine("         ,act_qty_jam_ke_18 ")
                                query.AppendLine("         ,act_amount_jam_ke_18 ")
                                query.AppendLine("         ,act_amount_jam_ke_18_shift ")
                                query.AppendLine("         ,act_amount_jam_ke_18_group ")

                                query.AppendLine("         ,act_qty_jam_ke_19 ")
                                query.AppendLine("         ,act_amount_jam_ke_19 ")
                                query.AppendLine("         ,act_amount_jam_ke_19_shift ")
                                query.AppendLine("         ,act_amount_jam_ke_19_group ")

                                query.AppendLine("         ,act_qty_jam_ke_20 ")
                                query.AppendLine("         ,act_amount_jam_ke_20 ")
                                query.AppendLine("         ,act_amount_jam_ke_20_shift ")
                                query.AppendLine("         ,act_amount_jam_ke_20_group ")

                                query.AppendLine("         ,act_qty_jam_ke_21 ")
                                query.AppendLine("         ,act_amount_jam_ke_21 ")
                                query.AppendLine("         ,act_amount_jam_ke_21_shift ")
                                query.AppendLine("         ,act_amount_jam_ke_21_group ")

                                query.AppendLine("         ,act_qty_jam_ke_22 ")
                                query.AppendLine("         ,act_amount_jam_ke_22 ")
                                query.AppendLine("         ,act_amount_jam_ke_22_shift ")
                                query.AppendLine("         ,act_amount_jam_ke_22_group ")

                                query.AppendLine("     ) ")
                                query.AppendLine(" values ")
                                query.AppendLine("     ( ")
                                query.AppendLine("         '" & str_date & "' ")
                                query.AppendLine("         ,'" & str_target_print_qty.ToString() & "' ")

                                query.AppendLine("         ,'" & str_target_amount_day.ToString() & "' ")
                                query.AppendLine("         ,'" & str_target_amount_akum.ToString() & "' ")
                                query.AppendLine("         ,'" & str_actual_amount_day.ToString() & "' ")
                                query.AppendLine("         ,'" & str_actual_amount_akum.ToString() & "' ")
                                query.AppendLine("         ,'" & str_dif_amount_target_actual.ToString() & "' ")

                                query.AppendLine("         ,'" & str_target_qty_day.ToString() & "' ")
                                query.AppendLine("         ,'" & str_target_qty_akum.ToString() & "' ")
                                query.AppendLine("         ,'" & str_actual_qty_day.ToString() & "' ")
                                query.AppendLine("         ,'" & str_actual_qty_akum.ToString() & "' ")
                                query.AppendLine("         ,'" & str_dif_qty_target_actual.ToString() & "' ")

                                query.AppendLine("         ,'" & str_average_size.ToString() & "' ")

                                query.AppendLine("         ,'" & act_qty_jam_ke_1.ToString() & "' ")
                                query.AppendLine("         ,'" & act_amount_jam_ke_1.ToString() & "' ")
                                query.AppendLine("         ,'" & act_amount_jam_ke_1_shift.ToString() & "' ")
                                query.AppendLine("         ,'" & act_amount_jam_ke_1_group.ToString() & "' ")

                                query.AppendLine("         ,'" & act_qty_jam_ke_2.ToString() & "' ")
                                query.AppendLine("         ,'" & act_amount_jam_ke_2.ToString() & "' ")
                                query.AppendLine("         ,'" & act_amount_jam_ke_2_shift.ToString() & "' ")
                                query.AppendLine("         ,'" & act_amount_jam_ke_2_group.ToString() & "' ")

                                query.AppendLine("         ,'" & act_qty_jam_ke_3.ToString() & "' ")
                                query.AppendLine("         ,'" & act_amount_jam_ke_3.ToString() & "' ")
                                query.AppendLine("         ,'" & act_amount_jam_ke_3_shift.ToString() & "' ")
                                query.AppendLine("         ,'" & act_amount_jam_ke_3_group.ToString() & "' ")

                                query.AppendLine("         ,'" & act_qty_jam_ke_4.ToString() & "' ")
                                query.AppendLine("         ,'" & act_amount_jam_ke_4.ToString() & "' ")
                                query.AppendLine("         ,'" & act_amount_jam_ke_4_shift.ToString() & "' ")
                                query.AppendLine("         ,'" & act_amount_jam_ke_4_group.ToString() & "' ")

                                query.AppendLine("         ,'" & act_qty_jam_ke_5.ToString() & "' ")
                                query.AppendLine("         ,'" & act_amount_jam_ke_5.ToString() & "' ")
                                query.AppendLine("         ,'" & act_amount_jam_ke_5_shift.ToString() & "' ")
                                query.AppendLine("         ,'" & act_amount_jam_ke_5_group.ToString() & "' ")

                                query.AppendLine("         ,'" & act_qty_jam_ke_6.ToString() & "' ")
                                query.AppendLine("         ,'" & act_amount_jam_ke_6.ToString() & "' ")
                                query.AppendLine("         ,'" & act_amount_jam_ke_6_shift.ToString() & "' ")
                                query.AppendLine("         ,'" & act_amount_jam_ke_6_group.ToString() & "' ")

                                query.AppendLine("         ,'" & act_qty_jam_ke_7.ToString() & "' ")
                                query.AppendLine("         ,'" & act_amount_jam_ke_7.ToString() & "' ")
                                query.AppendLine("         ,'" & act_amount_jam_ke_7_shift.ToString() & "' ")
                                query.AppendLine("         ,'" & act_amount_jam_ke_7_group.ToString() & "' ")

                                query.AppendLine("         ,'" & act_qty_jam_ke_8.ToString() & "' ")
                                query.AppendLine("         ,'" & act_amount_jam_ke_8.ToString() & "' ")
                                query.AppendLine("         ,'" & act_amount_jam_ke_8_shift.ToString() & "' ")
                                query.AppendLine("         ,'" & act_amount_jam_ke_8_group.ToString() & "' ")

                                query.AppendLine("         ,'" & act_qty_jam_ke_9.ToString() & "' ")
                                query.AppendLine("         ,'" & act_amount_jam_ke_9.ToString() & "' ")
                                query.AppendLine("         ,'" & act_amount_jam_ke_9_shift.ToString() & "' ")
                                query.AppendLine("         ,'" & act_amount_jam_ke_9_group.ToString() & "' ")

                                query.AppendLine("         ,'" & act_qty_jam_ke_10.ToString() & "' ")
                                query.AppendLine("         ,'" & act_amount_jam_ke_10.ToString() & "' ")
                                query.AppendLine("         ,'" & act_amount_jam_ke_10_shift.ToString() & "' ")
                                query.AppendLine("         ,'" & act_amount_jam_ke_10_group.ToString() & "' ")

                                query.AppendLine("         ,'" & act_qty_jam_ke_11.ToString() & "' ")
                                query.AppendLine("         ,'" & act_amount_jam_ke_11.ToString() & "' ")
                                query.AppendLine("         ,'" & act_amount_jam_ke_11_shift.ToString() & "' ")
                                query.AppendLine("         ,'" & act_amount_jam_ke_11_group.ToString() & "' ")

                                query.AppendLine("         ,'" & act_qty_jam_ke_12.ToString() & "' ")
                                query.AppendLine("         ,'" & act_amount_jam_ke_12.ToString() & "' ")
                                query.AppendLine("         ,'" & act_amount_jam_ke_12_shift.ToString() & "' ")
                                query.AppendLine("         ,'" & act_amount_jam_ke_12_group.ToString() & "' ")

                                query.AppendLine("         ,'" & act_qty_jam_ke_13.ToString() & "' ")
                                query.AppendLine("         ,'" & act_amount_jam_ke_13.ToString() & "' ")
                                query.AppendLine("         ,'" & act_amount_jam_ke_13_shift.ToString() & "' ")
                                query.AppendLine("         ,'" & act_amount_jam_ke_13_group.ToString() & "' ")

                                query.AppendLine("         ,'" & act_qty_jam_ke_14.ToString() & "' ")
                                query.AppendLine("         ,'" & act_amount_jam_ke_14.ToString() & "' ")
                                query.AppendLine("         ,'" & act_amount_jam_ke_14_shift.ToString() & "' ")
                                query.AppendLine("         ,'" & act_amount_jam_ke_14_group.ToString() & "' ")

                                query.AppendLine("         ,'" & act_qty_jam_ke_15_16_istirahat.ToString() & "' ")
                                query.AppendLine("         ,'" & act_amount_jam_ke_15_16_istirahat.ToString() & "' ")
                                query.AppendLine("         ,'" & act_amount_jam_ke_15_16_istirahat_shift.ToString() & "' ")
                                query.AppendLine("         ,'" & act_amount_jam_ke_15_16_istirahat_group.ToString() & "' ")

                                query.AppendLine("         ,'" & act_qty_jam_ke_17.ToString() & "' ")
                                query.AppendLine("         ,'" & act_amount_jam_ke_17.ToString() & "' ")
                                query.AppendLine("         ,'" & act_amount_jam_ke_17_shift.ToString() & "' ")
                                query.AppendLine("         ,'" & act_amount_jam_ke_17_group.ToString() & "' ")

                                query.AppendLine("         ,'" & act_qty_jam_ke_18.ToString() & "' ")
                                query.AppendLine("         ,'" & act_amount_jam_ke_18.ToString() & "' ")
                                query.AppendLine("         ,'" & act_amount_jam_ke_18_shift.ToString() & "' ")
                                query.AppendLine("         ,'" & act_amount_jam_ke_18_group.ToString() & "' ")

                                query.AppendLine("         ,'" & act_qty_jam_ke_19.ToString() & "' ")
                                query.AppendLine("         ,'" & act_amount_jam_ke_19.ToString() & "' ")
                                query.AppendLine("         ,'" & act_amount_jam_ke_19_shift.ToString() & "' ")
                                query.AppendLine("         ,'" & act_amount_jam_ke_19_group.ToString() & "' ")

                                query.AppendLine("         ,'" & act_qty_jam_ke_20.ToString() & "' ")
                                query.AppendLine("         ,'" & act_amount_jam_ke_20.ToString() & "' ")
                                query.AppendLine("         ,'" & act_amount_jam_ke_20_shift.ToString() & "' ")
                                query.AppendLine("         ,'" & act_amount_jam_ke_20_group.ToString() & "' ")

                                query.AppendLine("         ,'" & act_qty_jam_ke_21.ToString() & "' ")
                                query.AppendLine("         ,'" & act_amount_jam_ke_21.ToString() & "' ")
                                query.AppendLine("         ,'" & act_amount_jam_ke_21_shift.ToString() & "' ")
                                query.AppendLine("         ,'" & act_amount_jam_ke_21_group.ToString() & "' ")

                                query.AppendLine("         ,'" & act_qty_jam_ke_22.ToString() & "' ")
                                query.AppendLine("         ,'" & act_amount_jam_ke_22.ToString() & "' ")
                                query.AppendLine("         ,'" & act_amount_jam_ke_22_shift.ToString() & "' ")
                                query.AppendLine("         ,'" & act_amount_jam_ke_22_group.ToString() & "' ")

                                query.AppendLine("     ) ")
                                ClsConfig.ExecuteNonQuery(query.ToString(), ClsConfig.IPServer_ADDONS)
                                query.Length = 0
                                query.Capacity = 0
                                Console.WriteLine("Proses Insert Data")
                            End If

                        End If
                    Next

                    day = day.AddDays(1)
                End While

            End If

            query.AppendLine(" select ")
            query.AppendLine("     id ")
            query.AppendLine("     ,date ")
            query.AppendLine("     ,target_print_qty ")

            query.AppendLine("     ,target_amount_day ")
            query.AppendLine("     ,target_amount_akum ")
            query.AppendLine("     ,actual_amount_day ")
            query.AppendLine("     ,actual_amount_akum ")
            query.AppendLine("     ,dif_amount_target_actual ")

            query.AppendLine("     ,target_qty_day ")
            query.AppendLine("     ,target_qty_akum ")
            query.AppendLine("     ,actual_qty_day ")
            query.AppendLine("     ,actual_qty_akum ")
            query.AppendLine("     ,dif_qty_target_actual ")

            query.AppendLine("     ,average_size ")

            query.AppendLine("     ,act_qty_jam_ke_1 ")
            query.AppendLine("     ,act_amount_jam_ke_1 ")
            query.AppendLine("     ,act_qty_jam_ke_2 ")
            query.AppendLine("     ,act_amount_jam_ke_2 ")
            query.AppendLine("     ,act_qty_jam_ke_3 ")
            query.AppendLine("     ,act_amount_jam_ke_3 ")
            query.AppendLine("     ,act_qty_jam_ke_4 ")
            query.AppendLine("     ,act_amount_jam_ke_4 ")
            query.AppendLine("     ,act_qty_jam_ke_5 ")
            query.AppendLine("     ,act_amount_jam_ke_5 ")
            query.AppendLine("     ,act_qty_jam_ke_6 ")
            query.AppendLine("     ,act_amount_jam_ke_6 ")
            query.AppendLine("     ,act_qty_jam_ke_7 ")
            query.AppendLine("     ,act_amount_jam_ke_7 ")
            query.AppendLine("     ,act_qty_jam_ke_8 ")
            query.AppendLine("     ,act_amount_jam_ke_8 ")
            query.AppendLine("     ,act_qty_jam_ke_9 ")
            query.AppendLine("     ,act_amount_jam_ke_9 ")
            query.AppendLine("     ,act_qty_jam_ke_10 ")
            query.AppendLine("     ,act_amount_jam_ke_10 ")

            query.AppendLine("     ,act_qty_jam_ke_11 ")
            query.AppendLine("     ,act_amount_jam_ke_11 ")
            query.AppendLine("     ,act_qty_jam_ke_12 ")
            query.AppendLine("     ,act_amount_jam_ke_12 ")
            query.AppendLine("     ,act_qty_jam_ke_13 ")
            query.AppendLine("     ,act_amount_jam_ke_13 ")
            query.AppendLine("     ,act_qty_jam_ke_14 ")
            query.AppendLine("     ,act_amount_jam_ke_14 ")
            query.AppendLine("     ,act_qty_jam_ke_15_16_istirahat ")
            query.AppendLine("     ,act_amount_jam_ke_15_16_istirahat ")
            query.AppendLine("     ,act_qty_jam_ke_17 ")
            query.AppendLine("     ,act_amount_jam_ke_17 ")
            query.AppendLine("     ,act_qty_jam_ke_18 ")
            query.AppendLine("     ,act_amount_jam_ke_18 ")
            query.AppendLine("     ,act_qty_jam_ke_19 ")
            query.AppendLine("     ,act_amount_jam_ke_19 ")
            query.AppendLine("     ,act_qty_jam_ke_20 ")
            query.AppendLine("     ,act_amount_jam_ke_20 ")
            query.AppendLine("     ,act_qty_jam_ke_21 ")
            query.AppendLine("     ,act_amount_jam_ke_21 ")
            query.AppendLine("     ,act_qty_jam_ke_22 ")
            query.AppendLine("     ,act_amount_jam_ke_22 ")
            query.AppendLine(" from ")
            query.AppendLine("     ad_dis_rtjn_sum_qty_amount ")
            query.AppendLine(" where ")
            query.AppendLine("     date between '" + start_date.ToString("yyyy-MM-dd") + "' and '" + end_date.ToString("yyyy-MM-dd") + "' ")

            Dim dt_sum_result As DataTable
            dt_sum_result = ClsConfig.ExecuteQuery(query.ToString(), ClsConfig.IPServer_ADDONS)
            query.Length = 0
            query.Capacity = 0

            '(AMOUNT) EXPORT EXEL DAN SEND EMAIL
            If (str_status_email_amount = False) Then
                'JIKA PROGRAM GAGAL, EKPORT EXCEL DAN KIRIM ULANG
                export_excel_and_send_mail_amount(dt_sum_result, start_date, current_date, currentHHmm)
            Else
                'JIKA PROGRAM SUDAH LENGKAP PROSES, LANGSUNG DISKIP
                Console.WriteLine("PROGRAM SUDAH BERHASIL KALKULASI")
            End If

            '(QTY) EXPORT EXEL DAN SEND EMAIL
            If (str_status_email_qty = False) Then
                'JIKA PROGRAM GAGAL, EKPORT EXCEL DAN KIRIM ULANG
                export_excel_and_send_mail_qty(dt_sum_result, start_date, current_date, currentHHmm)
            Else
                'JIKA PROGRAM SUDAH LENGKAP PROSES, LANGSUNG DISKIP
                Console.WriteLine("PROGRAM SUDAH BERHASIL KALKULASI")
            End If

            dt_sum_result = Nothing

        Catch ex As Exception
            ClsConfig.create_log_error(currentDate, jenis_laporan, "[" + DateTime.Now.ToString("dd-MM-yyyy hh:mm:ss") + "] -- [ " + ex.Message + " ] -- Calculation Data Error")
            Environment.Exit(0)
        End Try

        Console.WriteLine("")
        Console.WriteLine("##FINISH CALCULATION")


        Console.WriteLine("")
        Console.WriteLine("##SUCCESS AND FINISH ALL PROCESSES")

    End Sub

    Private Sub export_excel_and_send_mail_amount(ByVal dt As DataTable, ByVal start_date As DateTime,
                                                  ByVal current_date As DateTime, ByVal currentHHmm As Integer)
        Console.WriteLine("")
        Console.WriteLine("##START EXPORT DATA TO EXCEL")
        Console.WriteLine("")

        Try

            Dim nama_file_template_n_path As String
            Dim nama_file_simpan As String
            Dim lokasi_simpan_file As String
            Dim mat_type As String = ""
            Dim OpenReport As Boolean = False 'Open file excel
            Dim ExcelOutputFile As String = ""

            'untuk jenis laporan
            jenis_laporan = "qty_amount"


            myCulture = New System.Globalization.CultureInfo("en-US", True)
            nama_file_template_n_path = System.AppDomain.CurrentDomain.BaseDirectory & ClsConfig.nama_file_template_amount & ".xlsx"
            'dir = System.Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
            lokasi_simpan_file = ClsConfig.lokasi_simpan_file_amount

            Dim xlApp As Object = CreateObject("Excel.Application")
            Dim xlWorkBook As Object = xlApp.Workbooks.Open(nama_file_template_n_path)
            Dim xlWorkSheet As Object

            xlWorkSheet = xlWorkBook.WorkSheets(1)
            Dim i As Integer
            Dim starting_row As Integer = 10 'awal data di tabel
            xlWorkSheet.Cells(1, 2) = "'" & start_date.ToString("dd-MM-yyyy") & " until " & current_date.ToString("dd-MM-yyyy")
            xlWorkSheet.Cells(2, 2) = "'" & Format(Now, "dd-MMM-yyyy HH:mm")

            Console.WriteLine("Total Export Record : " & dt.Rows.Count)
            For i = 0 To dt.Rows.Count - 1
                If dt(i)("id").ToString() <> "" Then
                    xlWorkSheet.Cells(4, 2) = dt(i)("target_print_qty").ToString()
                    xlWorkSheet.Cells(5, 2) = dt(i)("target_qty_day").ToString()
                    xlWorkSheet.Cells(i + starting_row, 2) = dt(i)("date").ToString()
                    xlWorkSheet.Cells(i + starting_row, 3) = dt(i)("date").ToString()
                    xlWorkSheet.Cells(i + starting_row, 4) = 1 'isi cell angka 1 untuk jumlah hari yang dihitung di template excel

                    xlWorkSheet.Cells(i + starting_row, 5) = dt(i)("target_amount_day").ToString()
                    'xlWorkSheet.Cells(i + starting_row, 6) = dt(i)("target_amount_akum").ToString() 'menggunakan rumus excel
                    xlWorkSheet.Cells(i + starting_row, 7) = dt(i)("actual_amount_day")
                    'xlWorkSheet.Cells(i + starting_row, 8) = dt(i)("actual_amount_akum").ToString() 'menggunakan rumus excel
                    'xlWorkSheet.Cells(i + starting_row, 9) = dt(i)("dif_amount_target_actual").ToString() 'menggunakan rumus excel

                    xlWorkSheet.Cells(i + starting_row, 10) = dt(i)("target_qty_day").ToString()
                    'xlWorkSheet.Cells(i + starting_row, 11) = dt(i)("target_qty_akum").ToString() 'menggunakan rumus excel
                    xlWorkSheet.Cells(i + starting_row, 12) = dt(i)("actual_qty_day").ToString()
                    'xlWorkSheet.Cells(i + starting_row, 13) = dt(i)("actual_qty_akum").ToString() 'menggunakan rumus excel
                    'xlWorkSheet.Cells(i + starting_row, 14) = dt(i)("dif_qty_target_actual").ToString() 'menggunakan rumus excel

                    xlWorkSheet.Cells(i + starting_row, 15) = dt(i)("average_size").ToString()

                    xlWorkSheet.Cells(i + starting_row, 16) = dt(i)("act_qty_jam_ke_1").ToString()
                    xlWorkSheet.Cells(i + starting_row, 17) = dt(i)("act_amount_jam_ke_1").ToString()
                    xlWorkSheet.Cells(i + starting_row, 18) = dt(i)("act_qty_jam_ke_2").ToString()
                    xlWorkSheet.Cells(i + starting_row, 19) = dt(i)("act_amount_jam_ke_2").ToString()
                    xlWorkSheet.Cells(i + starting_row, 20) = dt(i)("act_qty_jam_ke_3").ToString()
                    xlWorkSheet.Cells(i + starting_row, 21) = dt(i)("act_amount_jam_ke_3").ToString()
                    xlWorkSheet.Cells(i + starting_row, 22) = dt(i)("act_qty_jam_ke_4").ToString()
                    xlWorkSheet.Cells(i + starting_row, 23) = dt(i)("act_amount_jam_ke_4").ToString()
                    xlWorkSheet.Cells(i + starting_row, 24) = dt(i)("act_qty_jam_ke_5").ToString()
                    xlWorkSheet.Cells(i + starting_row, 25) = dt(i)("act_amount_jam_ke_5").ToString()

                    xlWorkSheet.Cells(i + starting_row, 26) = dt(i)("act_qty_jam_ke_6").ToString()
                    xlWorkSheet.Cells(i + starting_row, 27) = dt(i)("act_amount_jam_ke_6").ToString()
                    xlWorkSheet.Cells(i + starting_row, 28) = dt(i)("act_qty_jam_ke_7").ToString()
                    xlWorkSheet.Cells(i + starting_row, 29) = dt(i)("act_amount_jam_ke_7").ToString()
                    xlWorkSheet.Cells(i + starting_row, 30) = dt(i)("act_qty_jam_ke_8").ToString()
                    xlWorkSheet.Cells(i + starting_row, 31) = dt(i)("act_amount_jam_ke_8").ToString()
                    xlWorkSheet.Cells(i + starting_row, 32) = dt(i)("act_qty_jam_ke_9").ToString()
                    xlWorkSheet.Cells(i + starting_row, 33) = dt(i)("act_amount_jam_ke_9").ToString()
                    xlWorkSheet.Cells(i + starting_row, 34) = dt(i)("act_qty_jam_ke_10").ToString()
                    xlWorkSheet.Cells(i + starting_row, 35) = dt(i)("act_amount_jam_ke_10").ToString()

                    xlWorkSheet.Cells(i + starting_row, 36) = dt(i)("act_qty_jam_ke_11").ToString()
                    xlWorkSheet.Cells(i + starting_row, 37) = dt(i)("act_amount_jam_ke_11").ToString()
                    xlWorkSheet.Cells(i + starting_row, 38) = dt(i)("act_qty_jam_ke_12").ToString()
                    xlWorkSheet.Cells(i + starting_row, 39) = dt(i)("act_amount_jam_ke_12").ToString()
                    xlWorkSheet.Cells(i + starting_row, 40) = dt(i)("act_qty_jam_ke_13").ToString()
                    xlWorkSheet.Cells(i + starting_row, 41) = dt(i)("act_amount_jam_ke_13").ToString()
                    xlWorkSheet.Cells(i + starting_row, 42) = dt(i)("act_qty_jam_ke_14").ToString()
                    xlWorkSheet.Cells(i + starting_row, 43) = dt(i)("act_amount_jam_ke_14").ToString()
                    xlWorkSheet.Cells(i + starting_row, 44) = dt(i)("act_qty_jam_ke_15_16_istirahat").ToString()
                    xlWorkSheet.Cells(i + starting_row, 45) = dt(i)("act_amount_jam_ke_15_16_istirahat").ToString()

                    xlWorkSheet.Cells(i + starting_row, 46) = dt(i)("act_qty_jam_ke_17").ToString()
                    xlWorkSheet.Cells(i + starting_row, 47) = dt(i)("act_amount_jam_ke_17").ToString()
                    xlWorkSheet.Cells(i + starting_row, 48) = dt(i)("act_qty_jam_ke_18").ToString()
                    xlWorkSheet.Cells(i + starting_row, 49) = dt(i)("act_amount_jam_ke_18").ToString()
                    xlWorkSheet.Cells(i + starting_row, 50) = dt(i)("act_qty_jam_ke_19").ToString()
                    xlWorkSheet.Cells(i + starting_row, 51) = dt(i)("act_amount_jam_ke_19").ToString()
                    xlWorkSheet.Cells(i + starting_row, 52) = dt(i)("act_qty_jam_ke_20").ToString()
                    xlWorkSheet.Cells(i + starting_row, 53) = dt(i)("act_amount_jam_ke_20").ToString()
                    xlWorkSheet.Cells(i + starting_row, 54) = dt(i)("act_qty_jam_ke_21").ToString()
                    xlWorkSheet.Cells(i + starting_row, 55) = dt(i)("act_amount_jam_ke_21").ToString()
                    xlWorkSheet.Cells(i + starting_row, 56) = dt(i)("act_qty_jam_ke_22").ToString()
                    xlWorkSheet.Cells(i + starting_row, 57) = dt(i)("act_amount_jam_ke_22").ToString()
                    'xlWorkSheet.Range(xlWorkSheet.Cells(i + starting_row, j + 1), xlWorkSheet.Cells(i + starting_row, j + 1)).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red)
                End If
                Console.WriteLine("Export Record : " & i + 1)
            Next
            xlWorkSheet.Select()
            xlWorkSheet.Rows(57 + starting_row & ":1048576").Delete()
            xlWorkSheet.cells(1, 1).select()
            nama_file_simpan = ClsConfig.nama_file_lampiran_email_amount & "_" & Now.ToString("yyyyMMddHHmmss")
            xlWorkSheet.SaveAs(lokasi_simpan_file & "\" & nama_file_simpan & ".xlsx")
            xlWorkBook.Close()
            xlApp.Quit()
            releaseObject(xlApp)
            releaseObject(xlWorkBook)
            releaseObject(xlWorkSheet)
            ExcelOutputFile = lokasi_simpan_file & "\" & nama_file_simpan & ".xlsx"
            If OpenReport Then
                System.Diagnostics.Process.Start(lokasi_simpan_file & "\" & nama_file_simpan & ".xlsx")
            End If

            send_mail(ExcelOutputFile, dt, start_date, current_date, current_HHmm, "qty_amount")

        Catch ex As Exception
            ClsConfig.create_log_error(current_date, jenis_laporan, "[" + DateTime.Now.ToString("dd-MM-yyyy hh:mm:ss") + "] -- [ " + ex.Message + " ] -- Ekspor Data to Excel Error")
            Environment.Exit(0)
        End Try

        Console.WriteLine("")
        Console.WriteLine("##FINISH EKSPOR DATA")

    End Sub

    Private Sub export_excel_and_send_mail_qty(ByVal dt As DataTable, ByVal start_date As DateTime,
                                               ByVal current_date As DateTime, ByVal currentHHmm As Integer)
        Console.WriteLine("")
        Console.WriteLine("##START EXPORT DATA TO EXCEL")
        Console.WriteLine("")

        Try

            Dim nama_file_template_n_path As String
            Dim nama_file_simpan As String
            Dim lokasi_simpan_file As String
            Dim mat_type As String = ""
            Dim OpenReport As Boolean = False 'Open file excel
            Dim ExcelOutputFile As String = ""

            'untuk jenis report email
            jenis_laporan = "qty"

            myCulture = New System.Globalization.CultureInfo("en-US", True)
            nama_file_template_n_path = System.AppDomain.CurrentDomain.BaseDirectory & ClsConfig.nama_file_template_qty & ".xlsx"
            'dir = System.Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
            lokasi_simpan_file = ClsConfig.lokasi_simpan_file_qty

            Dim xlApp As Object = CreateObject("Excel.Application")
            Dim xlWorkBook As Object = xlApp.Workbooks.Open(nama_file_template_n_path)
            Dim xlWorkSheet As Object

            xlWorkSheet = xlWorkBook.WorkSheets(1)
            Dim i As Integer
            Dim starting_row As Integer = 9
            xlWorkSheet.Cells(1, 2) = "'" & start_date.ToString("dd-MM-yyyy") & " until " & current_date.ToString("dd-MM-yyyy")
            xlWorkSheet.Cells(2, 2) = "'" & Format(Now, "dd-MMM-yyyy HH:mm")

            Console.WriteLine("Total Export Record : " & dt.Rows.Count)
            For i = 0 To dt.Rows.Count - 1
                If dt(i)("id").ToString() <> "" Then
                    xlWorkSheet.Cells(4, 2) = dt(i)("target_print_qty").ToString()
                    xlWorkSheet.Cells(5, 2) = dt(i)("target_qty_day").ToString()

                    xlWorkSheet.Cells(i + starting_row, 2) = dt(i)("date").ToString()
                    xlWorkSheet.Cells(i + starting_row, 3) = dt(i)("date").ToString()
                    xlWorkSheet.Cells(i + starting_row, 4) = 1 'isi cell angka 1 untuk jumlah hari yang dihitung di template excel

                    xlWorkSheet.Cells(i + starting_row, 5) = dt(i)("target_qty_day").ToString()
                    'xlWorkSheet.Cells(i + starting_row, 6) = dt(i)("target_qty_akum").ToString() 'menggunakan rumus excel
                    xlWorkSheet.Cells(i + starting_row, 7) = dt(i)("actual_qty_day")
                    'xlWorkSheet.Cells(i + starting_row, 8) = dt(i)("actual_amount_akum").ToString() 'menggunakan rumus excel
                    'xlWorkSheet.Cells(i + starting_row, 9) = dt(i)("dif_amount_target_actual").ToString() 'menggunakan rumus excel

                    xlWorkSheet.Cells(i + starting_row, 10) = dt(i)("average_size").ToString()

                    xlWorkSheet.Cells(i + starting_row, 11) = dt(i)("act_qty_jam_ke_1").ToString()
                    xlWorkSheet.Cells(i + starting_row, 12) = dt(i)("act_qty_jam_ke_2").ToString()
                    xlWorkSheet.Cells(i + starting_row, 13) = dt(i)("act_qty_jam_ke_3").ToString()
                    xlWorkSheet.Cells(i + starting_row, 14) = dt(i)("act_qty_jam_ke_4").ToString()
                    xlWorkSheet.Cells(i + starting_row, 15) = dt(i)("act_qty_jam_ke_5").ToString()
                    xlWorkSheet.Cells(i + starting_row, 16) = dt(i)("act_qty_jam_ke_6").ToString()
                    xlWorkSheet.Cells(i + starting_row, 17) = dt(i)("act_qty_jam_ke_7").ToString()
                    xlWorkSheet.Cells(i + starting_row, 18) = dt(i)("act_qty_jam_ke_8").ToString()
                    xlWorkSheet.Cells(i + starting_row, 19) = dt(i)("act_qty_jam_ke_9").ToString()
                    xlWorkSheet.Cells(i + starting_row, 20) = dt(i)("act_qty_jam_ke_10").ToString()

                    xlWorkSheet.Cells(i + starting_row, 21) = dt(i)("act_qty_jam_ke_11").ToString()
                    xlWorkSheet.Cells(i + starting_row, 22) = dt(i)("act_qty_jam_ke_12").ToString()
                    xlWorkSheet.Cells(i + starting_row, 23) = dt(i)("act_qty_jam_ke_13").ToString()
                    xlWorkSheet.Cells(i + starting_row, 24) = dt(i)("act_qty_jam_ke_14").ToString()
                    xlWorkSheet.Cells(i + starting_row, 25) = dt(i)("act_qty_jam_ke_15_16_istirahat").ToString()
                    xlWorkSheet.Cells(i + starting_row, 26) = dt(i)("act_qty_jam_ke_17").ToString()
                    xlWorkSheet.Cells(i + starting_row, 27) = dt(i)("act_qty_jam_ke_18").ToString()
                    xlWorkSheet.Cells(i + starting_row, 28) = dt(i)("act_qty_jam_ke_19").ToString()
                    xlWorkSheet.Cells(i + starting_row, 29) = dt(i)("act_qty_jam_ke_20").ToString()
                    xlWorkSheet.Cells(i + starting_row, 30) = dt(i)("act_qty_jam_ke_21").ToString()
                    xlWorkSheet.Cells(i + starting_row, 31) = dt(i)("act_qty_jam_ke_22").ToString()
                    'xlWorkSheet.Range(xlWorkSheet.Cells(i + starting_row, j + 1), xlWorkSheet.Cells(i + starting_row, j + 1)).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red)
                End If
                Console.WriteLine("Export Record : " & i + 1)
            Next
            xlWorkSheet.Select()
            xlWorkSheet.Rows(31 + starting_row & ":1048576").Delete()
            xlWorkSheet.cells(1, 1).select()
            nama_file_simpan = ClsConfig.nama_file_lampiran_email_qty & "_" & Now.ToString("yyyyMMddHHmmss")
            xlWorkSheet.SaveAs(lokasi_simpan_file & "\" & nama_file_simpan & ".xlsx")
            xlWorkBook.Close()
            xlApp.Quit()
            releaseObject(xlApp)
            releaseObject(xlWorkBook)
            releaseObject(xlWorkSheet)
            ExcelOutputFile = lokasi_simpan_file & "\" & nama_file_simpan & ".xlsx"
            If OpenReport Then
                System.Diagnostics.Process.Start(lokasi_simpan_file & "\" & nama_file_simpan & ".xlsx")
            End If

            send_mail(ExcelOutputFile, dt, start_date, current_date, currentHHmm, "qty")

        Catch ex As Exception
            ClsConfig.create_log_error(current_date, jenis_laporan, "[" + DateTime.Now.ToString("dd-MM-yyyy hh:mm:ss") + "] -- [ " + ex.Message + " ] -- Ekspor Data to Excel Error")
            Environment.Exit(0)
        End Try

        Console.WriteLine("")
        Console.WriteLine("##FINISH EKSPOR DATA")

    End Sub

    Private Sub releaseObject(ByVal obj As Object)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            obj = Nothing
        Catch ex As Exception
            obj = Nothing
        Finally
            GC.Collect()
        End Try
    End Sub

    Private Sub send_mail(ByVal AttachedFile As String, ByVal dtSource As DataTable, ByVal start_date As DateTime,
                          ByVal current_date As DateTime, ByVal currentHHmm As Integer, ByVal JenisLaporan As String)
        Console.WriteLine("")
        Console.WriteLine("##START CREATE AND SEND MAIL")

        Try

            If (JenisLaporan = "qty") Then
                subject_email = ClsConfig.subject_email_monitoring_qty
            ElseIf (JenisLaporan = "qty_amount") Then
                subject_email = ClsConfig.subject_email_monitoring_amount
            End If

            If AttachedFile = "" Then Exit Sub
            Dim tbTemp As New DataTable
            Dim query As StringBuilder = New StringBuilder()

            Dim AddressMail_To As String = ""
            Dim body_message As New StringBuilder

            If JenisLaporan = "qty_amount" Then
                If Not create_body_msg_amount(body_message, dtSource, start_date, current_date) Then Exit Sub

                query.AppendLine(" SELECT MAILADDRESS FROM Z_TANTO_LIST Where QTY_AMOUNT_RTJN IN ('To') ORDER BY Asc_Email_Sort DESC ")
                'query.AppendLine(" SELECT MAILADDRESS FROM Z_TANTO_LIST Where QTY_AMOUNT_RTJN IN ('To') AND name in('Halim') ORDER BY Asc_Email_Sort DESC ")
            Else
                If Not create_body_msg_qty(body_message, dtSource, start_date, current_date) Then Exit Sub

                query.AppendLine(" SELECT MAILADDRESS FROM Z_TANTO_LIST Where QTY_RTJN IN ('To') ORDER BY Asc_Email_Sort DESC ")
                'query.AppendLine(" SELECT MAILADDRESS FROM Z_TANTO_LIST Where QTY_RTJN IN ('To') AND name in('Halim') ORDER BY Asc_Email_Sort DESC ")
            End If

            tbTemp = ClsConfig.ExecuteQuery(query.ToString(), ClsConfig.IPServer_TxDTIPRD)
            query.Length = 0
            query.Capacity = 0

            If tbTemp.Rows.Count > 0 Then
                Dim vw As DataView = tbTemp.DefaultView
                Dim tb As Data.DataTable = vw.ToTable()
                Dim rdr As DataTableReader = tb.CreateDataReader()
                While rdr.Read
                    AddressMail_To = rdr("MAILADDRESS") & "," & AddressMail_To
                End While
                rdr.Close()
            End If

            If Microsoft.VisualBasic.Right(Trim(AddressMail_To), 1) = "," Then AddressMail_To = Microsoft.VisualBasic.Left(AddressMail_To, Len(AddressMail_To) - 1)

            SendExcelMailViaSMTP(AddressMail_To, body_message, AttachedFile, start_date, current_date, currentHHmm, JenisLaporan)

        Catch ex As Exception
            'Panggil fungsi send email
            send_mail(AttachedFile, dtSource, start_date, current_date, currentHHmm, JenisLaporan)

            ClsConfig.create_log_error(current_date, JenisLaporan, "[" + DateTime.Now.ToString("dd-MM-yyyy hh:mm:ss") + "] -- [ " + ex.Message + " ] -- Create Email Error")
            Environment.Exit(0)
        End Try

        Console.WriteLine("##FINISH CREATE AND SEND MAIL")
    End Sub

    Private Function create_body_msg_amount(ByRef body_str As StringBuilder, ByRef dtSource As DataTable, ByRef start_date As DateTime, ByRef current_date As DateTime) As Boolean

        Dim Result As Boolean = False
        Dim misValue As Object = Missing.Value
        Dim body_str_temp As New StringBuilder

        If dtSource.Rows.Count > 0 Then
            Result = True
            body_str_temp.AppendLine("<html>")
            body_str_temp.AppendLine("<body>")
            body_str_temp.AppendLine("Dear All, <br />")
            body_str_temp.AppendLine("<br />")
            body_str_temp.AppendLine("This is Qty (RTJN) and Amount (TPiCS) by period : " & start_date.ToString("dd-MMM-yyyy") & " until " & current_date.ToString("dd-MMM-yyyy") & " <br />")
            body_str_temp.AppendLine("Please find the attached file for detailed information <br /><br />")
            body_str_temp.AppendLine("<table style='border-collapse: collapse'>")

            body_str_temp.AppendLine("<tr>")
            body_str_temp.AppendLine("<td rowspan='2' style='text-align: left; font-size:12px; border: 1px solid black; padding-right:10px; padding-left:4px;' > " & " Date " & "</td>")
            body_str_temp.AppendLine("<td colspan='5' style='text-align: center; font-size:12px; border: 1px solid black; padding-right:10px; padding-left:4px;' > " & " Amount " & "</td>")
            body_str_temp.AppendLine("<td colspan='5' style='text-align: center; font-size:12px; border: 1px solid black; padding-right:10px; padding-left:4px;' > " & " Pcs " & "</td>")
            body_str_temp.AppendLine("<td rowspan='2' style='text-align: left; font-size:12px; border: 1px solid black; padding-right:10px; padding-left:4px;' > " & " Ave inch " & "</td>")
            body_str_temp.AppendLine("</tr>")

            body_str_temp.AppendLine("<tr>")
            body_str_temp.AppendLine("<td style='text-align: left; font-size:12px; border: 1px solid black; padding-right:10px; padding-left:4px;' > " & " Target (Day) " & "</td>")
            body_str_temp.AppendLine("<td style='text-align: left; font-size:12px; border: 1px solid black; padding-right:10px; padding-left:4px;' > " & " Target (Amount) " & "</td>")
            body_str_temp.AppendLine("<td style='text-align: left; font-size:12px; border: 1px solid black; padding-right:10px; padding-left:4px;' > " & " Actual (Day) " & "</td>")
            body_str_temp.AppendLine("<td style='text-align: left; font-size:12px; border: 1px solid black; padding-right:10px; padding-left:4px;' > " & " Actual (Amount) " & "</td>")
            body_str_temp.AppendLine("<td style='text-align: left; font-size:12px; border: 1px solid black; padding-right:10px; padding-left:4px;' > " & " Diff " & "</td>")

            body_str_temp.AppendLine("<td style='text-align: left; font-size:12px; border: 1px solid black; padding-right:10px; padding-left:4px;' > " & " Target (Day) " & "</td>")
            body_str_temp.AppendLine("<td style='text-align: left; font-size:12px; border: 1px solid black; padding-right:10px; padding-left:4px;' > " & " Target (Pcs) " & "</td>")
            body_str_temp.AppendLine("<td style='text-align: left; font-size:12px; border: 1px solid black; padding-right:10px; padding-left:4px;' > " & " Actual (Day) " & "</td>")
            body_str_temp.AppendLine("<td style='text-align: left; font-size:12px; border: 1px solid black; padding-right:10px; padding-left:4px;' > " & " Actual (Pcs) " & "</td>")
            body_str_temp.AppendLine("<td style='text-align: left; font-size:12px; border: 1px solid black; padding-right:10px; padding-left:4px;' > " & " Diff " & "</td>")
            body_str_temp.AppendLine("</tr>")

            For i = 0 To dtSource.Rows.Count - 1
                body_str_temp.AppendLine("<tr>")
                body_str_temp.AppendLine("<td style='text-align: left; font-size:12px; border: 1px solid black; padding-right:10px; padding-left:4px;' >" & DateTime.Parse(dtSource(i)("date")).ToString("dd-MM-yyyy") & "</td>")

                body_str_temp.AppendLine("<td style='text-align: left; font-size:12px; border: 1px solid black; padding-right:10px; padding-left:4px;' >" & Double.Parse(dtSource(i)("target_amount_day")).ToString("#,##0") & "</td>")
                body_str_temp.AppendLine("<td style='text-align: left; font-size:12px; border: 1px solid black; padding-right:10px; padding-left:4px;' >" & Double.Parse(dtSource(i)("target_amount_akum")).ToString("#,##0") & "</td>")
                body_str_temp.AppendLine("<td style='text-align: left; font-size:12px; border: 1px solid black; padding-right:10px; padding-left:4px;' >" & Double.Parse(dtSource(i)("actual_amount_day")).ToString("#,##0") & "</td>")
                body_str_temp.AppendLine("<td style='text-align: left; font-size:12px; border: 1px solid black; padding-right:10px; padding-left:4px;' >" & Double.Parse(dtSource(i)("actual_amount_akum")).ToString("#,##0") & "</td>")
                body_str_temp.AppendLine("<td style='text-align: left; font-size:12px; border: 1px solid black; padding-right:10px; padding-left:4px;' >" & Double.Parse(dtSource(i)("dif_amount_target_actual")).ToString("#,##0") & "</td>")

                body_str_temp.AppendLine("<td style='text-align: left; font-size:12px; border: 1px solid black; padding-right:10px; padding-left:4px;' >" & Double.Parse(dtSource(i)("target_qty_day")).ToString("#,##0") & "</td>")
                body_str_temp.AppendLine("<td style='text-align: left; font-size:12px; border: 1px solid black; padding-right:10px; padding-left:4px;' >" & Double.Parse(dtSource(i)("target_qty_akum")).ToString("#,##0") & "</td>")
                body_str_temp.AppendLine("<td style='text-align: left; font-size:12px; border: 1px solid black; padding-right:10px; padding-left:4px;' >" & Double.Parse(dtSource(i)("actual_qty_day")).ToString("#,##0") & "</td>")
                body_str_temp.AppendLine("<td style='text-align: left; font-size:12px; border: 1px solid black; padding-right:10px; padding-left:4px;' >" & Double.Parse(dtSource(i)("actual_qty_akum")).ToString("#,##0") & "</td>")
                body_str_temp.AppendLine("<td style='text-align: left; font-size:12px; border: 1px solid black; padding-right:10px; padding-left:4px;' >" & Double.Parse(dtSource(i)("dif_qty_target_actual")).ToString("#,##0") & "</td>")

                body_str_temp.AppendLine("<td style='text-align: left; font-size:12px; border: 1px solid black; padding-right:10px; padding-left:4px;' >" & Double.Parse(dtSource(i)("average_size")).ToString("0.0") & "</td>")
                body_str_temp.AppendLine("</tr>")
            Next
            body_str_temp.AppendLine("</table>")
            body_str_temp.AppendLine("</body>")
            body_str_temp.AppendLine("</html>")
        Else
            Result = False
        End If
        body_str = body_str_temp
        create_body_msg_amount = Result
    End Function

    Private Function create_body_msg_qty(ByRef body_str As StringBuilder, ByRef dtSource As DataTable, ByRef start_date As DateTime, ByRef current_date As DateTime) As Boolean

        Dim Result As Boolean = False
        Dim misValue As Object = Missing.Value
        Dim body_str_temp As New StringBuilder

        If dtSource.Rows.Count > 0 Then
            Result = True
            body_str_temp.AppendLine("<html>")
            body_str_temp.AppendLine("<body>")
            body_str_temp.AppendLine("Dear All, <br />")
            body_str_temp.AppendLine("<br />")
            body_str_temp.AppendLine("This is Qty (RTJN) by period : " & start_date.ToString("dd-MMM-yyyy") & " until " & current_date.ToString("dd-MMM-yyyy") & " <br />")
            body_str_temp.AppendLine("Please find the attached file for detailed information <br /><br />")
            body_str_temp.AppendLine("<table style='border-collapse: collapse'>")

            body_str_temp.AppendLine("<tr>")
            body_str_temp.AppendLine("<td style='text-align: left; font-size:12px; border: 1px solid black; padding-right:10px; padding-left:4px;' > " & " Date " & "</td>")

            body_str_temp.AppendLine("<td style='text-align: left; font-size:12px; border: 1px solid black; padding-right:10px; padding-left:4px;' > " & " Target (Day) " & "</td>")
            body_str_temp.AppendLine("<td style='text-align: left; font-size:12px; border: 1px solid black; padding-right:10px; padding-left:4px;' > " & " Target (Pcs) " & "</td>")
            body_str_temp.AppendLine("<td style='text-align: left; font-size:12px; border: 1px solid black; padding-right:10px; padding-left:4px;' > " & " Actual (Day) " & "</td>")
            body_str_temp.AppendLine("<td style='text-align: left; font-size:12px; border: 1px solid black; padding-right:10px; padding-left:4px;' > " & " Actual (Pcs) " & "</td>")
            body_str_temp.AppendLine("<td style='text-align: left; font-size:12px; border: 1px solid black; padding-right:10px; padding-left:4px;' > " & " Diff " & "</td>")

            body_str_temp.AppendLine("<td style='text-align: left; font-size:12px; border: 1px solid black; padding-right:10px; padding-left:4px;' > " & " Ave inch " & "</td>")
            body_str_temp.AppendLine("</tr>")

            For i = 0 To dtSource.Rows.Count - 1
                body_str_temp.AppendLine("<tr>")
                body_str_temp.AppendLine("<td style='text-align: left; font-size:12px; border: 1px solid black; padding-right:10px; padding-left:4px;' >" & DateTime.Parse(dtSource(i)("date")).ToString("dd-MM-yyyy") & "</td>")
                body_str_temp.AppendLine("<td style='text-align: left; font-size:12px; border: 1px solid black; padding-right:10px; padding-left:4px;' >" & Double.Parse(dtSource(i)("target_qty_day")).ToString("#,##0") & "</td>")
                body_str_temp.AppendLine("<td style='text-align: left; font-size:12px; border: 1px solid black; padding-right:10px; padding-left:4px;' >" & Double.Parse(dtSource(i)("target_qty_akum")).ToString("#,##0") & "</td>")
                body_str_temp.AppendLine("<td style='text-align: left; font-size:12px; border: 1px solid black; padding-right:10px; padding-left:4px;' >" & Double.Parse(dtSource(i)("actual_qty_day")).ToString("#,##0") & "</td>")
                body_str_temp.AppendLine("<td style='text-align: left; font-size:12px; border: 1px solid black; padding-right:10px; padding-left:4px;' >" & Double.Parse(dtSource(i)("actual_qty_akum")).ToString("#,##0") & "</td>")
                body_str_temp.AppendLine("<td style='text-align: left; font-size:12px; border: 1px solid black; padding-right:10px; padding-left:4px;' >" & Double.Parse(dtSource(i)("dif_qty_target_actual")).ToString("#,##0") & "</td>")

                body_str_temp.AppendLine("<td style='text-align: left; font-size:12px; border: 1px solid black; padding-right:10px; padding-left:4px;' >" & Double.Parse(dtSource(i)("average_size")).ToString("0.0") & "</td>")
                body_str_temp.AppendLine("</tr>")
            Next
            body_str_temp.AppendLine("</table>")
            body_str_temp.AppendLine("</body>")
            body_str_temp.AppendLine("</html>")
        Else
            Result = False
        End If
        body_str = body_str_temp
        create_body_msg_qty = Result
    End Function


    Private Function SendExcelMailViaSMTP(
                                            ByVal strToAddress As String,
                                            ByVal BodyMsg As StringBuilder,
                                            ByVal AttachedFile As String,
                                            ByVal start_date As DateTime,
                                            ByVal current_date As DateTime,
                                            ByVal currentHHmm As Integer,
                                            ByVal JenisLaporan As String
                                          ) As Boolean

        Dim query As StringBuilder = New StringBuilder()
        Dim email_nama As String = ClsConfig.email_nama
        Dim email_password As String = ClsConfig.email_password
        Dim email_server_smtp As String = ClsConfig.email_server_smtp
        Dim email_server_port As String = ClsConfig.email_server_port
        Dim subject_email As String = ClsConfig.subject_email
        'Dim tls_1_2 = DirectCast(3072, System.Net.SecurityProtocolType) 'TLS 1.2 //old
        Dim tls As Int32 = ClsConfig.tls 'Get tls from .ini
        Dim tls_1_2 = DirectCast(tls, System.Net.SecurityProtocolType) 'TLS 1.2
        'Dim date_now As String = Format(Now)

        Dim oMail As New MailMessage()
        Dim oSmtp As New SmtpClient
        oSmtp.UseDefaultCredentials = False
        oSmtp.Credentials = New Net.NetworkCredential(email_nama, email_password)
        oSmtp.Port = CInt(email_server_port)
        oSmtp.EnableSsl = True
        oSmtp.Host = email_server_smtp

        oMail = New MailMessage()
        oMail.From = New MailAddress(email_nama)
        oMail.To.Add(strToAddress)
        oMail.Subject = subject_email & " : " & start_date.ToString("dd-MMM-yyyy") & " until " & current_date.ToString("dd-MMM-yyyy")
        oMail.IsBodyHtml = True
        oMail.Body = BodyMsg.ToString
        oMail.Attachments.Add(New Attachment(AttachedFile))
        System.Net.ServicePointManager.Expect100Continue = False
        System.Net.ServicePointManager.SecurityProtocol = tls_1_2

        Dim message As String = ""

        'SEND EMAIL
        Try
            status_sudah_email = ClsGnrl.cek_status_sudah_email(current_date, currentHHmm, JenisLaporan)
            If (status_sudah_email = True) Then
                'JIKA PROGRAM SUDAH KALKULASI DAN KIRIM EMAIL, DI-SKIP PROSES
                Console.WriteLine("PROGRAM SUDAH KALKULASI DAN KIRIM EMAIL")
            Else
                oSmtp.Send(oMail)
                ClsGnrl.monitoring_email(current_date, current_HHmm, JenisLaporan, message)
            End If
        Catch ex As Exception
            'Panggil fungsi send email agar kirim email ulang.
            SendExcelMailViaSMTP(strToAddress, BodyMsg, AttachedFile, start_date, current_date, currentHHmm, JenisLaporan)

            ClsConfig.create_log_error(current_date, JenisLaporan, "[" + DateTime.Now.ToString("dd-MM-yyyy hh:mm:ss") + "] -- [ " + ex.Message + " ] -- Send Email Error")
            Environment.Exit(0)
        End Try

        If JenisLaporan = "qty_amount" Then
            query.AppendLine(" update ")
            query.AppendLine("     ad_dis_rtjn_sum_qty_amount ")
            query.AppendLine(" set ")
            query.AppendLine("     status_email_amount='sudah kirim' ")
            query.AppendLine("     ,datetime_email_amount=getdate() ")
            query.AppendLine(" where ")
            query.AppendLine("     date = '" + current_date.ToString("yyyy-MM-dd") + "' ")
        Else
            query.AppendLine(" update ")
            query.AppendLine("     ad_dis_rtjn_sum_qty_amount ")
            query.AppendLine(" set ")
            query.AppendLine("     status_email_qty='sudah kirim' ")
            query.AppendLine("     ,datetime_email_qty=getdate() ")
            query.AppendLine(" where ")
            query.AppendLine("     date = '" + current_date.ToString("yyyy-MM-dd") + "' ")
        End If

        ClsConfig.ExecuteNonQuery(query.ToString(), ClsConfig.IPServer_ADDONS)
        query.Length = 0
        query.Capacity = 0
        Console.WriteLine("--> Update Status Email: " & JenisLaporan)

    End Function

End Module
