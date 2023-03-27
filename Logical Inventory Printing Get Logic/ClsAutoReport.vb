Imports System.Text 'untuk library teks query yang panjang
Imports System.Net.Mail 'untuk library email

Public Class ClsAutoReport

    Dim ClsGnrl As New ClsGeneral
    Dim NamaTable_Monitoring As String
    Dim NamaReport As String
    Dim jenis_mail_sender_report As String
    Dim jenis_mail_sender_report1 As String
    Dim jenis_mail_sender_report2 As String

    'Reporting Var
    Dim nama_file_template_n_path As String
    Dim nama_file_simpan As String
    Dim lokasi_simpan_file As String

    'Config Email Var
    Dim email_nama As String
    Dim email_password As String
    Dim email_server_smtp As String
    Dim email_server_port As String
    Dim subject_email As String
    Dim tbTemp As New DataTable
    Dim query As StringBuilder = New StringBuilder()

    Dim AddressMail_To As String
    Dim body_message As New StringBuilder

    'Log Error Var
    Dim keterangan_error As String


    Public Sub AutoReportMonitoring(ByVal startDate As DateTime, ByVal currentDate As DateTime, ByVal jenis_laporan As String)
        Try
            'Table Monitoring Maintenance
            NamaTable_Monitoring = "ad_dis_monitoring_maintenance"

            If (jenis_laporan = "qty") Then
                jenis_mail_sender_report = "Realtime Production (Qty)"
            ElseIf (jenis_laporan = "qty_amount") Then
                jenis_mail_sender_report = "Realtime Production (Amount)"
            End If

            Dim query As New StringBuilder
            Dim dt As DataTable

            query.AppendLine(" select ")
            query.AppendLine("     date ")
            query.AppendLine("     ,pic_follow_up ")
            query.AppendLine("     ,[Pukul_08_30] ")
            query.AppendLine("     ,[Pukul_09_30] ")
            query.AppendLine("     ,[Pukul_10_30] ")
            query.AppendLine("     ,[Pukul_11_30] ")
            query.AppendLine("     ,[Pukul_12_30] ")
            query.AppendLine("     ,[Pukul_13_30] ")
            query.AppendLine("     ,[Pukul_14_30] ")
            query.AppendLine("     ,[Pukul_15_30] ")
            query.AppendLine("     ,[Pukul_16_30] ")
            query.AppendLine("     ,[Pukul_17_30] ")
            query.AppendLine("     ,[Pukul_18_30] ")
            query.AppendLine("     ,[Pukul_19_30] ")
            query.AppendLine("     ,[Pukul_20_30] ")
            query.AppendLine("     ,[Pukul_21_30] ")
            query.AppendLine("     ,[Pukul_22_30] ")
            query.AppendLine("     ,[Pukul_23_30] ")
            query.AppendLine("     ,[Pukul_00_30] ")
            query.AppendLine("     ,[Pukul_01_30] ")
            query.AppendLine("     ,[Pukul_02_30] ")
            query.AppendLine("     ,[Pukul_03_30] ")
            query.AppendLine("     ,[Pukul_04_30] ")
            query.AppendLine("     ,[Pukul_05_30] ")
            query.AppendLine("     ,[Pukul_06_30] ")
            query.AppendLine("     ,[Pukul_07_30] ")
            query.AppendLine("     ,jumlah_kegagalan ")
            query.AppendLine("     ,masalah_kegagalan ")
            query.AppendLine("     ,aksi_solusi_perbaikan ")
            query.AppendLine("     ,status_perbaikan ")
            query.AppendLine("     ,keterangan_laporan ")
            query.AppendLine("     ,last_email_sent ")
            query.AppendLine(" from ")
            query.AppendLine("     " & NamaTable_Monitoring & " ")
            query.AppendLine(" where ")
            query.AppendLine("     jenis_mail_sender like '%" & jenis_mail_sender_report & "%' ")
            query.AppendLine("     and FORMAT(date, 'yyyyMM') = '" & currentDate.ToString("yyyyMM") & "' ")
            dt = ClsConfig.ExecuteQuery(query.ToString(), ClsConfig.IPServer_ADDONS)
            query.Length = 0
            query.Capacity = 0

            If (jenis_laporan = "qty") Then
                nama_file_template_n_path = System.AppDomain.CurrentDomain.BaseDirectory & ClsConfig.nama_file_template_monitoring_qty & ".xlsx"
                nama_file_simpan = ClsConfig.nama_file_lampiran_email_monitoring_qty & "_" & Now.ToString("yyyyMMddHHmmss")
                lokasi_simpan_file = ClsConfig.lokasi_simpan_file_monitoring_qty

                'Create File Excel and Send
                CreateExcelFile(dt, startDate, currentDate, jenis_laporan, nama_file_template_n_path, nama_file_simpan, lokasi_simpan_file, jenis_mail_sender_report)

            ElseIf (jenis_laporan = "qty_amount") Then
                nama_file_template_n_path = System.AppDomain.CurrentDomain.BaseDirectory & ClsConfig.nama_file_template_monitoring_amount & ".xlsx"
                nama_file_simpan = ClsConfig.nama_file_lampiran_email_monitoring_amount & "_" & Now.ToString("yyyyMMddHHmmss")
                lokasi_simpan_file = ClsConfig.lokasi_simpan_file_monitoring_amount

                'Create File Excel and Send
                CreateExcelFile(dt, startDate, currentDate, jenis_laporan, nama_file_template_n_path, nama_file_simpan, lokasi_simpan_file, jenis_mail_sender_report)

            End If

        Catch ex As Exception
            ClsConfig.create_log_error(currentDate, jenis_laporan, "[" + DateTime.Now.ToString("dd-MM-yyyy hh:mm:ss") + "] -- [ " + ex.Message + " ] --Export excel monitoring mail sender Realtime Production")
            Environment.Exit(0)
        End Try

    End Sub
    Private Sub CreateExcelFile(
                          ByVal dt As DataTable,
                          ByVal startDate As DateTime,
                          ByVal currentDate As DateTime,
                          ByVal jenis_laporan As String,
                          ByVal nama_file_template_n_path As String,
                          ByVal nama_file_simpan As String,
                          ByVal lokasi_simpan_file As String,
                          ByVal jenis_mail_sender_report As String
                          )

        Dim ExcelOutputFile As String = ""

        Dim xlApp As Object = CreateObject("Excel.Application")
        Dim xlWorkBook As Object = xlApp.Workbooks.Open(nama_file_template_n_path)
        Dim xlWorkSheet1 As Object
        Dim starting_row As Integer
        Dim row_count As Integer
        Dim last_row As Integer

        xlWorkSheet1 = xlWorkBook.WorkSheets(1)
        starting_row = 7
        row_count = dt.Rows.Count
        last_row = row_count + starting_row

        xlWorkSheet1.Cells(1, 1) = jenis_mail_sender_report 'berdasarkan jenis mail sender
        xlWorkSheet1.Cells(2, 1) = Format(startDate, "dd-MMM-yyyy") & " until " & Format(currentDate, "dd-MMM-yyyy")
        xlWorkSheet1.Cells(3, 1) = "Printed date : " & Format(Now, "dd-MMM-yyyy HH:mm")

        For i = 0 To row_count - 1
            If dt(i)("date").ToString() <> "" Then
                xlWorkSheet1.Cells(i + starting_row, 1) = (i + 1)
                xlWorkSheet1.Cells(i + starting_row, 2) = dt(i)("date").ToString()
                xlWorkSheet1.Cells(i + starting_row, 3) = dt(i)("pic_follow_up").ToString()
                xlWorkSheet1.Cells(i + starting_row, 4) = If(dt(i)("Pukul_08_30").ToString() <> "", dt(i)("Pukul_08_30").ToString(), "NG")
                xlWorkSheet1.Cells(i + starting_row, 5) = If(dt(i)("Pukul_09_30").ToString() <> "", dt(i)("Pukul_09_30").ToString(), "NG")
                xlWorkSheet1.Cells(i + starting_row, 6) = If(dt(i)("Pukul_10_30").ToString() <> "", dt(i)("Pukul_10_30").ToString(), "NG")
                xlWorkSheet1.Cells(i + starting_row, 7) = If(dt(i)("Pukul_11_30").ToString() <> "", dt(i)("Pukul_11_30").ToString(), "NG")
                xlWorkSheet1.Cells(i + starting_row, 8) = If(dt(i)("Pukul_12_30").ToString() <> "", dt(i)("Pukul_12_30").ToString(), "NG")
                xlWorkSheet1.Cells(i + starting_row, 9) = If(dt(i)("Pukul_13_30").ToString() <> "", dt(i)("Pukul_13_30").ToString(), "NG")
                xlWorkSheet1.Cells(i + starting_row, 10) = If(dt(i)("Pukul_14_30").ToString() <> "", dt(i)("Pukul_14_30").ToString(), "NG")
                xlWorkSheet1.Cells(i + starting_row, 11) = If(dt(i)("Pukul_15_30").ToString() <> "", dt(i)("Pukul_15_30").ToString(), "NG")
                xlWorkSheet1.Cells(i + starting_row, 12) = If(dt(i)("Pukul_16_30").ToString() <> "", dt(i)("Pukul_16_30").ToString(), "NG")
                xlWorkSheet1.Cells(i + starting_row, 13) = If(dt(i)("Pukul_17_30").ToString() <> "", dt(i)("Pukul_17_30").ToString(), "NG")
                xlWorkSheet1.Cells(i + starting_row, 14) = If(dt(i)("Pukul_18_30").ToString() <> "", dt(i)("Pukul_18_30").ToString(), "NG")
                xlWorkSheet1.Cells(i + starting_row, 15) = If(dt(i)("Pukul_19_30").ToString() <> "", dt(i)("Pukul_19_30").ToString(), "NG")
                xlWorkSheet1.Cells(i + starting_row, 16) = If(dt(i)("Pukul_20_30").ToString() <> "", dt(i)("Pukul_20_30").ToString(), "NG")
                xlWorkSheet1.Cells(i + starting_row, 17) = If(dt(i)("Pukul_21_30").ToString() <> "", dt(i)("Pukul_21_30").ToString(), "NG")
                xlWorkSheet1.Cells(i + starting_row, 18) = If(dt(i)("Pukul_22_30").ToString() <> "", dt(i)("Pukul_22_30").ToString(), "NG")
                xlWorkSheet1.Cells(i + starting_row, 19) = If(dt(i)("Pukul_23_30").ToString() <> "", dt(i)("Pukul_23_30").ToString(), "NG")
                xlWorkSheet1.Cells(i + starting_row, 20) = If(dt(i)("Pukul_00_30").ToString() <> "", dt(i)("Pukul_00_30").ToString(), "NG")
                xlWorkSheet1.Cells(i + starting_row, 21) = If(dt(i)("Pukul_01_30").ToString() <> "", dt(i)("Pukul_01_30").ToString(), "NG")
                xlWorkSheet1.Cells(i + starting_row, 22) = If(dt(i)("Pukul_02_30").ToString() <> "", dt(i)("Pukul_02_30").ToString(), "NG")
                xlWorkSheet1.Cells(i + starting_row, 23) = If(dt(i)("Pukul_03_30").ToString() <> "", dt(i)("Pukul_03_30").ToString(), "NG")
                xlWorkSheet1.Cells(i + starting_row, 24) = If(dt(i)("Pukul_04_30").ToString() <> "", dt(i)("Pukul_04_30").ToString(), "NG")
                xlWorkSheet1.Cells(i + starting_row, 25) = If(dt(i)("Pukul_05_30").ToString() <> "", dt(i)("Pukul_05_30").ToString(), "NG")
                xlWorkSheet1.Cells(i + starting_row, 26) = If(dt(i)("Pukul_06_30").ToString() <> "", dt(i)("Pukul_06_30").ToString(), "NG")
                xlWorkSheet1.Cells(i + starting_row, 27) = If(dt(i)("Pukul_07_30").ToString() <> "", dt(i)("Pukul_07_30").ToString(), "NG")
                xlWorkSheet1.Cells(i + starting_row, 30) = dt(i)("masalah_kegagalan").ToString()
                xlWorkSheet1.Cells(i + starting_row, 31) = dt(i)("aksi_solusi_perbaikan").ToString()
                xlWorkSheet1.Cells(i + starting_row, 32) = dt(i)("status_perbaikan").ToString()
                xlWorkSheet1.Cells(i + starting_row, 33) = dt(i)("keterangan_laporan").ToString()
            End If
        Next

        xlWorkSheet1.Select()
        xlWorkSheet1.Rows(last_row & ":1048576").Delete()
        xlWorkSheet1.cells(1, 1).select()

        'xlWorkSheet1.SaveAs(lokasi_simpan_file & "\" & nama_file_simpan & ".xlsx") 'simpan hanya 1 sheet
        xlApp.ActiveWorkbook.SaveAs(lokasi_simpan_file & "\" & nama_file_simpan & ".xlsx") 'simpan beberapa sheet

        xlApp.Quit()
        releaseObject(xlApp)
        releaseObject(xlWorkBook)
        releaseObject(xlWorkSheet1)

        ExcelOutputFile = lokasi_simpan_file & "\" & nama_file_simpan & ".xlsx"

        Threading.Thread.Sleep(5000)

        'Send Email
        send_mail(ExcelOutputFile, dt, startDate, currentDate, jenis_laporan, jenis_mail_sender_report)
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

    Private Sub send_mail(ByVal AttachedFile As String,
                          ByVal dtSource As DataTable,
                          ByVal start_date As DateTime,
                          ByVal current_date As DateTime,
                          ByVal jenis_laporan As String,
                          ByVal jenis_mail_sender_report As String
                          )

        Try
            If AttachedFile = "" Then Exit Sub

            'reset AddressMail
            AddressMail_To = ""

            If (jenis_laporan = "qty") Then
                query.AppendLine(" SELECT MAILADDRESS FROM Z_TANTO_LIST Where MAILADDRESS in ('','" & ClsConfig.email_monitoring_mail_sender_qty & "') ORDER BY Asc_Email_Sort DESC ")
            ElseIf (jenis_laporan = "qty_amount") Then
                query.AppendLine(" SELECT MAILADDRESS FROM Z_TANTO_LIST Where MAILADDRESS in ('','" & ClsConfig.email_monitoring_mail_sender_amount & "') ORDER BY Asc_Email_Sort DESC ")

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

            If Not create_body_msg(body_message, dtSource, start_date, current_date, jenis_laporan, jenis_mail_sender_report) Then Exit Sub

            SendExcelMailViaSMTP(AddressMail_To, body_message, AttachedFile, start_date, current_date, jenis_laporan, jenis_mail_sender_report)

        Catch ex As Exception
            'Panggil fungsi send email agar kirim email ulang.
            send_mail(AttachedFile, dtSource, start_date, current_date, jenis_laporan, jenis_mail_sender_report)

            ClsConfig.create_log_error(current_date, jenis_laporan, "[" + DateTime.Now.ToString("dd-MM-yyyy hh:mm:ss") + "] -- [ " + ex.Message + " ] -- Send email monitoring mail sender loss amount error")
            Environment.Exit(0)
        End Try
    End Sub

    Private Function create_body_msg(ByRef body_str As StringBuilder,
                                     ByRef dtSource As DataTable,
                                     ByRef startDate As DateTime,
                                     ByRef currentDate As DateTime,
                                     ByVal jenis_laporan As String,
                                     ByVal jenis_mail_sender_report As String) As Boolean
        Try
            Dim Result As Boolean = False
            Dim body_str_temp As New StringBuilder

            If dtSource.Rows.Count > 0 Then
                Result = True
                body_str_temp.AppendLine("<html>")
                body_str_temp.AppendLine("<body>")
                body_str_temp.AppendLine("Dear All, <br />")
                body_str_temp.AppendLine("This is the Daily Monitoring Mail Sender " & jenis_mail_sender_report & " Report by period : " &
                                         Format(startDate, "dd-MMM-yyyy") & " to " & Format(currentDate, "dd-MMM-yyyy") &
                                         " <br />")
                body_str_temp.AppendLine("Please find the attached file for detailed information <br /><br />")
                body_str_temp.AppendLine("<table style='border-collapse: collapse'>")

                body_str_temp.AppendLine("<tr>")
                body_str_temp.AppendLine("<td  style='text-align: center; font-weight: bold; font-size:12px; border: 1px solid black; padding-right:10px; padding-left:4px;' >" & " NO " & "  " & "</td>")
                body_str_temp.AppendLine("<td  style='text-align: center; font-weight: bold; font-size:12px; border: 1px solid black; padding-right:10px; padding-left:4px;' >" & " TANGGAL " & "  " & "</td>")
                body_str_temp.AppendLine("<td  style='text-align: center; font-weight: bold; font-size:12px; border: 1px solid black; padding-right:10px; padding-left:4px;' >" & " 08:30 " & "  " & "</td>")
                body_str_temp.AppendLine("<td  style='text-align: center; font-weight: bold; font-size:12px; border: 1px solid black; padding-right:10px; padding-left:4px;' >" & " 09:30 " & "  " & "</td>")
                body_str_temp.AppendLine("<td  style='text-align: center; font-weight: bold; font-size:12px; border: 1px solid black; padding-right:10px; padding-left:4px;' >" & " 10:30 " & "  " & "</td>")
                body_str_temp.AppendLine("<td  style='text-align: center; font-weight: bold; font-size:12px; border: 1px solid black; padding-right:10px; padding-left:4px;' >" & " 11:30 " & "  " & "</td>")
                body_str_temp.AppendLine("<td  style='text-align: center; font-weight: bold; font-size:12px; border: 1px solid black; padding-right:10px; padding-left:4px;' >" & " 12:30 " & "  " & "</td>")
                body_str_temp.AppendLine("<td  style='text-align: center; font-weight: bold; font-size:12px; border: 1px solid black; padding-right:10px; padding-left:4px;' >" & " 13:30 " & "  " & "</td>")
                body_str_temp.AppendLine("<td  style='text-align: center; font-weight: bold; font-size:12px; border: 1px solid black; padding-right:10px; padding-left:4px;' >" & " 14:30 " & "  " & "</td>")
                body_str_temp.AppendLine("<td  style='text-align: center; font-weight: bold; font-size:12px; border: 1px solid black; padding-right:10px; padding-left:4px;' >" & " 15:30 " & "  " & "</td>")
                body_str_temp.AppendLine("<td  style='text-align: center; font-weight: bold; font-size:12px; border: 1px solid black; padding-right:10px; padding-left:4px;' >" & " 16:30 " & "  " & "</td>")
                body_str_temp.AppendLine("<td  style='text-align: center; font-weight: bold; font-size:12px; border: 1px solid black; padding-right:10px; padding-left:4px;' >" & " 17:30 " & "  " & "</td>")
                body_str_temp.AppendLine("<td  style='text-align: center; font-weight: bold; font-size:12px; border: 1px solid black; padding-right:10px; padding-left:4px;' >" & " 18:30 " & "  " & "</td>")
                body_str_temp.AppendLine("<td  style='text-align: center; font-weight: bold; font-size:12px; border: 1px solid black; padding-right:10px; padding-left:4px;' >" & " 19:30 " & "  " & "</td>")
                body_str_temp.AppendLine("<td  style='text-align: center; font-weight: bold; font-size:12px; border: 1px solid black; padding-right:10px; padding-left:4px;' >" & " 20:30 " & "  " & "</td>")
                body_str_temp.AppendLine("<td  style='text-align: center; font-weight: bold; font-size:12px; border: 1px solid black; padding-right:10px; padding-left:4px;' >" & " 21:30 " & "  " & "</td>")
                body_str_temp.AppendLine("<td  style='text-align: center; font-weight: bold; font-size:12px; border: 1px solid black; padding-right:10px; padding-left:4px;' >" & " 22:30 " & "  " & "</td>")
                body_str_temp.AppendLine("<td  style='text-align: center; font-weight: bold; font-size:12px; border: 1px solid black; padding-right:10px; padding-left:4px;' >" & " 23:30 " & "  " & "</td>")
                body_str_temp.AppendLine("<td  style='text-align: center; font-weight: bold; font-size:12px; border: 1px solid black; padding-right:10px; padding-left:4px;' >" & " 00:30 " & "  " & "</td>")
                body_str_temp.AppendLine("<td  style='text-align: center; font-weight: bold; font-size:12px; border: 1px solid black; padding-right:10px; padding-left:4px;' >" & " 01:30 " & "  " & "</td>")
                body_str_temp.AppendLine("<td  style='text-align: center; font-weight: bold; font-size:12px; border: 1px solid black; padding-right:10px; padding-left:4px;' >" & " 02:30 " & "  " & "</td>")
                body_str_temp.AppendLine("<td  style='text-align: center; font-weight: bold; font-size:12px; border: 1px solid black; padding-right:10px; padding-left:4px;' >" & " 03:30 " & "  " & "</td>")
                body_str_temp.AppendLine("<td  style='text-align: center; font-weight: bold; font-size:12px; border: 1px solid black; padding-right:10px; padding-left:4px;' >" & " 04:30 " & "  " & "</td>")
                body_str_temp.AppendLine("<td  style='text-align: center; font-weight: bold; font-size:12px; border: 1px solid black; padding-right:10px; padding-left:4px;' >" & " 05:30 " & "  " & "</td>")
                body_str_temp.AppendLine("<td  style='text-align: center; font-weight: bold; font-size:12px; border: 1px solid black; padding-right:10px; padding-left:4px;' >" & " 06:30 " & "  " & "</td>")
                body_str_temp.AppendLine("<td  style='text-align: center; font-weight: bold; font-size:12px; border: 1px solid black; padding-right:10px; padding-left:4px;' >" & " 07:30 " & "  " & "</td>")
                body_str_temp.AppendLine("<td  style='text-align: center; font-weight: bold; font-size:12px; border: 1px solid black; padding-right:10px; padding-left:4px;' >" & " JUMLAH KEGAGALAN " & "  " & "</td>")
                body_str_temp.AppendLine("</tr>")

                For i = 0 To dtSource.Rows.Count - 1
                    body_str_temp.AppendLine("<tr>")
                    body_str_temp.AppendLine("<td  style='text-align: center; font-size:12px; border: 1px solid black; padding-right:10px; padding-left:4px;' >" & (i + 1).ToString() & "</td>")
                    body_str_temp.AppendLine("<td  style='text-align: center; font-size:12px; border: 1px solid black; padding-right:10px; padding-left:4px;' >" & DateTime.Parse(dtSource(i)("date")).ToString("dd-MM-yyyy") & "</td>")
                    body_str_temp.AppendLine("<td  style='text-align: center; font-size:12px; border: 1px solid black; padding-right:10px; padding-left:4px;" & If(dtSource(i)("Pukul_08_30").ToString() <> "", "", " background-color:red;") & "' >" & If(dtSource(i)("Pukul_08_30").ToString() <> "", dtSource(i)("Pukul_08_30").ToString(), "NG") & "</td>")
                    body_str_temp.AppendLine("<td  style='text-align: center; font-size:12px; border: 1px solid black; padding-right:10px; padding-left:4px;" & If(dtSource(i)("Pukul_09_30").ToString() <> "", "", " background-color:red;") & "' >" & If(dtSource(i)("Pukul_09_30").ToString() <> "", dtSource(i)("Pukul_09_30").ToString(), "NG") & "</td>")
                    body_str_temp.AppendLine("<td  style='text-align: center; font-size:12px; border: 1px solid black; padding-right:10px; padding-left:4px;" & If(dtSource(i)("Pukul_10_30").ToString() <> "", "", " background-color:red;") & "' >" & If(dtSource(i)("Pukul_10_30").ToString() <> "", dtSource(i)("Pukul_10_30").ToString(), "NG") & "</td>")
                    body_str_temp.AppendLine("<td  style='text-align: center; font-size:12px; border: 1px solid black; padding-right:10px; padding-left:4px;" & If(dtSource(i)("Pukul_11_30").ToString() <> "", "", " background-color:red;") & "' >" & If(dtSource(i)("Pukul_11_30").ToString() <> "", dtSource(i)("Pukul_11_30").ToString(), "NG") & "</td>")
                    body_str_temp.AppendLine("<td  style='text-align: center; font-size:12px; border: 1px solid black; padding-right:10px; padding-left:4px;" & If(dtSource(i)("Pukul_12_30").ToString() <> "", "", " background-color:red;") & "' >" & If(dtSource(i)("Pukul_12_30").ToString() <> "", dtSource(i)("Pukul_12_30").ToString(), "NG") & "</td>")
                    body_str_temp.AppendLine("<td  style='text-align: center; font-size:12px; border: 1px solid black; padding-right:10px; padding-left:4px;" & If(dtSource(i)("Pukul_13_30").ToString() <> "", "", " background-color:red;") & "' >" & If(dtSource(i)("Pukul_13_30").ToString() <> "", dtSource(i)("Pukul_13_30").ToString(), "NG") & "</td>")
                    body_str_temp.AppendLine("<td  style='text-align: center; font-size:12px; border: 1px solid black; padding-right:10px; padding-left:4px;" & If(dtSource(i)("Pukul_14_30").ToString() <> "", "", " background-color:red;") & "' >" & If(dtSource(i)("Pukul_14_30").ToString() <> "", dtSource(i)("Pukul_14_30").ToString(), "NG") & "</td>")
                    body_str_temp.AppendLine("<td  style='text-align: center; font-size:12px; border: 1px solid black; padding-right:10px; padding-left:4px;" & If(dtSource(i)("Pukul_15_30").ToString() <> "", "", " background-color:red;") & "' >" & If(dtSource(i)("Pukul_15_30").ToString() <> "", dtSource(i)("Pukul_15_30").ToString(), "NG") & "</td>")
                    body_str_temp.AppendLine("<td  style='text-align: center; font-size:12px; border: 1px solid black; padding-right:10px; padding-left:4px;" & If(dtSource(i)("Pukul_16_30").ToString() <> "", "", " background-color:red;") & "' >" & If(dtSource(i)("Pukul_16_30").ToString() <> "", dtSource(i)("Pukul_16_30").ToString(), "NG") & "</td>")
                    body_str_temp.AppendLine("<td  style='text-align: center; font-size:12px; border: 1px solid black; padding-right:10px; padding-left:4px;" & If(dtSource(i)("Pukul_17_30").ToString() <> "", "", " background-color:red;") & "' >" & If(dtSource(i)("Pukul_17_30").ToString() <> "", dtSource(i)("Pukul_17_30").ToString(), "NG") & "</td>")
                    body_str_temp.AppendLine("<td  style='text-align: center; font-size:12px; border: 1px solid black; padding-right:10px; padding-left:4px;" & If(dtSource(i)("Pukul_18_30").ToString() <> "", "", " background-color:red;") & "' >" & If(dtSource(i)("Pukul_18_30").ToString() <> "", dtSource(i)("Pukul_18_30").ToString(), "NG") & "</td>")
                    body_str_temp.AppendLine("<td  style='text-align: center; font-size:12px; border: 1px solid black; padding-right:10px; padding-left:4px;" & If(dtSource(i)("Pukul_19_30").ToString() <> "", "", " background-color:red;") & "' >" & If(dtSource(i)("Pukul_19_30").ToString() <> "", dtSource(i)("Pukul_19_30").ToString(), "NG") & "</td>")
                    body_str_temp.AppendLine("<td  style='text-align: center; font-size:12px; border: 1px solid black; padding-right:10px; padding-left:4px;" & If(dtSource(i)("Pukul_20_30").ToString() <> "", "", " background-color:red;") & "' >" & If(dtSource(i)("Pukul_20_30").ToString() <> "", dtSource(i)("Pukul_20_30").ToString(), "NG") & "</td>")
                    body_str_temp.AppendLine("<td  style='text-align: center; font-size:12px; border: 1px solid black; padding-right:10px; padding-left:4px;" & If(dtSource(i)("Pukul_21_30").ToString() <> "", "", " background-color:red;") & "' >" & If(dtSource(i)("Pukul_21_30").ToString() <> "", dtSource(i)("Pukul_21_30").ToString(), "NG") & "</td>")
                    body_str_temp.AppendLine("<td  style='text-align: center; font-size:12px; border: 1px solid black; padding-right:10px; padding-left:4px;" & If(dtSource(i)("Pukul_22_30").ToString() <> "", "", " background-color:red;") & "' >" & If(dtSource(i)("Pukul_22_30").ToString() <> "", dtSource(i)("Pukul_22_30").ToString(), "NG") & "</td>")
                    body_str_temp.AppendLine("<td  style='text-align: center; font-size:12px; border: 1px solid black; padding-right:10px; padding-left:4px;" & If(dtSource(i)("Pukul_23_30").ToString() <> "", "", " background-color:red;") & "' >" & If(dtSource(i)("Pukul_23_30").ToString() <> "", dtSource(i)("Pukul_23_30").ToString(), "NG") & "</td>")
                    body_str_temp.AppendLine("<td  style='text-align: center; font-size:12px; border: 1px solid black; padding-right:10px; padding-left:4px;" & If(dtSource(i)("Pukul_00_30").ToString() <> "", "", " background-color:red;") & "' >" & If(dtSource(i)("Pukul_00_30").ToString() <> "", dtSource(i)("Pukul_00_30").ToString(), "NG") & "</td>")
                    body_str_temp.AppendLine("<td  style='text-align: center; font-size:12px; border: 1px solid black; padding-right:10px; padding-left:4px;" & If(dtSource(i)("Pukul_01_30").ToString() <> "", "", " background-color:red;") & "' >" & If(dtSource(i)("Pukul_01_30").ToString() <> "", dtSource(i)("Pukul_01_30").ToString(), "NG") & "</td>")
                    body_str_temp.AppendLine("<td  style='text-align: center; font-size:12px; border: 1px solid black; padding-right:10px; padding-left:4px;" & If(dtSource(i)("Pukul_02_30").ToString() <> "", "", " background-color:red;") & "' >" & If(dtSource(i)("Pukul_02_30").ToString() <> "", dtSource(i)("Pukul_02_30").ToString(), "NG") & "</td>")
                    body_str_temp.AppendLine("<td  style='text-align: center; font-size:12px; border: 1px solid black; padding-right:10px; padding-left:4px;" & If(dtSource(i)("Pukul_03_30").ToString() <> "", "", " background-color:red;") & "' >" & If(dtSource(i)("Pukul_03_30").ToString() <> "", dtSource(i)("Pukul_03_30").ToString(), "NG") & "</td>")
                    body_str_temp.AppendLine("<td  style='text-align: center; font-size:12px; border: 1px solid black; padding-right:10px; padding-left:4px;" & If(dtSource(i)("Pukul_04_30").ToString() <> "", "", " background-color:red;") & "' >" & If(dtSource(i)("Pukul_04_30").ToString() <> "", dtSource(i)("Pukul_04_30").ToString(), "NG") & "</td>")
                    body_str_temp.AppendLine("<td  style='text-align: center; font-size:12px; border: 1px solid black; padding-right:10px; padding-left:4px;" & If(dtSource(i)("Pukul_05_30").ToString() <> "", "", " background-color:red;") & "' >" & If(dtSource(i)("Pukul_05_30").ToString() <> "", dtSource(i)("Pukul_05_30").ToString(), "NG") & "</td>")
                    body_str_temp.AppendLine("<td  style='text-align: center; font-size:12px; border: 1px solid black; padding-right:10px; padding-left:4px;" & If(dtSource(i)("Pukul_06_30").ToString() <> "", "", " background-color:red;") & "' >" & If(dtSource(i)("Pukul_06_30").ToString() <> "", dtSource(i)("Pukul_06_30").ToString(), "NG") & "</td>")
                    body_str_temp.AppendLine("<td  style='text-align: center; font-size:12px; border: 1px solid black; padding-right:10px; padding-left:4px;" & If(dtSource(i)("Pukul_07_30").ToString() <> "", "", " background-color:red;") & "' >" & If(dtSource(i)("Pukul_07_30").ToString() <> "", dtSource(i)("Pukul_07_30").ToString(), "NG") & "</td>")
                    If i <> (dtSource.Rows.Count - 1) Then
                        body_str_temp.AppendLine("<td  style='text-align: center; font-size:12px; border: 1px solid black; padding-right:10px; padding-left:4px;' >" & dtSource(i)("jumlah_kegagalan").ToString() & "</td>")
                    Else
                        body_str_temp.AppendLine("<td  style='text-align: center; font-size:12px; border: 1px solid black; padding-right:10px; padding-left:4px;' >" & (Convert.ToInt32(dtSource(i)("jumlah_kegagalan").ToString()) - 1).ToString() & "</td>")  'ditambah penjumlahan -1 (minus 1) karena 1 nilai belum update, karena statusnya baru akan update
                    End If
                    body_str_temp.AppendLine("</tr>")
                Next
                body_str_temp.AppendLine("</table>")

                body_str_temp.AppendLine("</body>")
                body_str_temp.AppendLine("</html>")
            Else
                Result = False
            End If
            body_str = body_str_temp
            create_body_msg = Result

        Catch ex As Exception
            ClsConfig.create_log_error(currentDate, jenis_laporan, "[" + DateTime.Now.ToString("dd-MM-yyyy hh:mm:ss") + "] -- [ " + ex.Message + " ] -- Create content/body email monitoring mail sender loss amount error")
            Environment.Exit(0)
        End Try
    End Function

    Private Function SendExcelMailViaSMTP(
                                            ByVal strToAddress As String,
                                            ByVal BodyMsg As StringBuilder,
                                            ByVal AttachedFile As String,
                                            ByVal start_date As DateTime,
                                            ByVal current_date As DateTime,
                                            ByVal jenis_laporan As String,
                                            ByVal jenis_mail_sender_report As String
                                          ) As Boolean

        Try

            Dim query As StringBuilder = New StringBuilder()
            email_nama = ClsConfig.email_nama
            email_password = ClsConfig.email_password
            email_server_smtp = ClsConfig.email_server_smtp
            email_server_port = ClsConfig.email_server_port

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

            If (jenis_laporan = "qty") Then
                subject_email = ClsConfig.subject_email_monitoring_qty
            Else
                subject_email = ClsConfig.subject_email_monitoring_amount
            End If

            oMail.Subject = subject_email & " : " & start_date.ToString("dd-MMM-yyyy") & " until " & current_date.ToString("dd-MMM-yyyy")
            oMail.IsBodyHtml = True
            oMail.Body = BodyMsg.ToString
            oMail.Attachments.Add(New Attachment(AttachedFile))
            System.Net.ServicePointManager.Expect100Continue = False
            System.Net.ServicePointManager.SecurityProtocol = tls_1_2

            Dim message As String = "" 'isi kosong jika tidak ada error

            Console.WriteLine("----> PROSES KIRIM FILE MONITORING: " + jenis_laporan)

            'SEND EMAIL
            oSmtp.Send(oMail)

            ClsGnrl.monitoring_email(current_date, 9999, jenis_laporan, message)

        Catch ex As Exception
            'Panggil fungsi send email agar kirim email ulang.
            SendExcelMailViaSMTP(strToAddress, BodyMsg, AttachedFile, start_date, current_date, jenis_laporan, jenis_mail_sender_report)

            ClsConfig.create_log_error(current_date, jenis_laporan, "[" + DateTime.Now.ToString("dd-MM-yyyy hh:mm:ss") + "] -- [ " + ex.Message + " ] -- Send Email in SMTP Error")
            Environment.Exit(0)
        End Try

    End Function

End Class
