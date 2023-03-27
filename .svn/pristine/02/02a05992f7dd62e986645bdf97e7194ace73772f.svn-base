Imports Microsoft.VisualBasic
Imports System.Data.SqlClient
Imports System.Text
Imports System.Runtime.InteropServices
Imports System.IO

Public Class ClsConfig
    Public Shared SQL As String
    Public Shared Cn As SqlConnection
    Public Shared Cmd As SqlCommand
    Public Shared Da As SqlDataAdapter
    Public Shared Ds As DataSet
    Public Shared Dt As DataTable


    Public Shared DATABASE_TYPE As String
    Public Shared IPServer_RTJN_PRD As String
    Public Shared IPServer_TxDTIPRD As String
    Public Shared IPServer_ADDONS As String

    Public Shared email_from_alias As String
    Public Shared email_nama As String
    Public Shared email_password As String
    Public Shared email_server_smtp As String
    Public Shared email_server_port As String
    Public Shared subject_email As String
    Public Shared tls As String

    Public Shared nama_folder_log_error As String
    Public Shared nama_file_txt_log_error As String

    Public Shared nama_file_template_qty As String
    Public Shared nama_file_template_amount As String
    Public Shared nama_file_lampiran_email_qty As String
    Public Shared nama_file_lampiran_email_amount As String
    Public Shared lokasi_simpan_file_qty As String
    Public Shared lokasi_simpan_file_amount As String

    'monitoring qty 
    Public Shared nama_file_template_monitoring_qty As String
    Public Shared nama_file_lampiran_email_monitoring_qty As String
    Public Shared lokasi_simpan_file_monitoring_qty As String
    Public Shared subject_email_monitoring_qty As String
    Public Shared email_monitoring_mail_sender_qty As String

    'monitoring amount
    Public Shared nama_file_template_monitoring_amount As String
    Public Shared nama_file_lampiran_email_monitoring_amount As String
    Public Shared lokasi_simpan_file_monitoring_amount As String
    Public Shared subject_email_monitoring_amount As String
    Public Shared email_monitoring_mail_sender_amount As String

    'report
    Dim ClsAutRep As New ClsAutoReport
    Dim start_date, end_date, current_date As DateTime
    Dim current_HHmm As Integer 'jam dan menit yang sedang berjalan




    <DllImport("kernel32.dll")>
    Private Shared Function GetPrivateProfileString(ByVal lpApplicationName As String,
                                                    ByVal lpKeyName As String,
                                                    ByVal lpDefault As String,
                                                    ByVal lpReturnedString As StringBuilder,
                                                    ByVal nSize As UInt32,
                                                    ByVal lpFileName As String) As UInt32
    End Function

    Private Shared Function GetIniString(ByVal iniFileName As String,
                                 ByVal section As String,
                                 ByVal key As String,
                                 Optional ByVal defaultValue As String = "") As String
        Dim nSize As Integer = 1024
        Dim sb As StringBuilder = New StringBuilder(nSize)
        Dim ret As UInt32 = GetPrivateProfileString(section, key, defaultValue, sb, Convert.ToUInt32(sb.Capacity), iniFileName)

        Return sb.ToString
    End Function

    Public Shared Sub get_variable_setting()
        Dim EXE_PATH As String
        'server database
        EXE_PATH = System.AppDomain.CurrentDomain.BaseDirectory
        DATABASE_TYPE = GetIniString(EXE_PATH & "\Auto_RTJN_Qty_Amount.ini", "DATABASE", "TYPE")
        IPServer_RTJN_PRD = GetIniString(EXE_PATH & "\Auto_RTJN_Qty_Amount.ini", "DATABASE", "RTJN")
        IPServer_TxDTIPRD = GetIniString(EXE_PATH & "\Auto_RTJN_Qty_Amount.ini", "DATABASE", "TPICS")
        IPServer_ADDONS = GetIniString(EXE_PATH & "\Auto_RTJN_Qty_Amount.ini", "DATABASE", "ADDONS")

        'email
        email_from_alias = GetIniString(EXE_PATH & "\Auto_RTJN_Qty_Amount.ini", "EMAIL", "email_from_alias")
        email_nama = GetIniString(EXE_PATH & "\Auto_RTJN_Qty_Amount.ini", "EMAIL", "email_nama")
        email_password = GetIniString(EXE_PATH & "\Auto_RTJN_Qty_Amount.ini", "EMAIL", "email_password")
        email_server_smtp = GetIniString(EXE_PATH & "\Auto_RTJN_Qty_Amount.ini", "EMAIL", "email_server_smtp")
        email_server_port = GetIniString(EXE_PATH & "\Auto_RTJN_Qty_Amount.ini", "EMAIL", "email_server_port")
        subject_email = GetIniString(EXE_PATH & "\Auto_RTJN_Qty_Amount.ini", "EMAIL", "subject_email")
        tls = GetIniString(EXE_PATH & "\Auto_RTJN_Qty_Amount.ini", "EMAIL", "tls")

        'history error
        nama_folder_log_error = GetIniString(EXE_PATH & "\Auto_RTJN_Qty_Amount.ini", "FILE", "nama_folder_log_error")
        nama_file_txt_log_error = GetIniString(EXE_PATH & "\Auto_RTJN_Qty_Amount.ini", "FILE", "nama_file_txt_log_error")

        'export excel email
        nama_file_template_qty = GetIniString(EXE_PATH & "\Auto_RTJN_Qty_Amount.ini", "FILE", "nama_file_template_qty")
        nama_file_template_amount = GetIniString(EXE_PATH & "\Auto_RTJN_Qty_Amount.ini", "FILE", "nama_file_template_amount")
        nama_file_lampiran_email_qty = GetIniString(EXE_PATH & "\Auto_RTJN_Qty_Amount.ini", "FILE", "nama_file_lampiran_email_qty")
        nama_file_lampiran_email_amount = GetIniString(EXE_PATH & "\Auto_RTJN_Qty_Amount.ini", "FILE", "nama_file_lampiran_email_amount")
        lokasi_simpan_file_qty = GetIniString(EXE_PATH & "\Auto_RTJN_Qty_Amount.ini", "FILE", "lokasi_simpan_file_qty")
        lokasi_simpan_file_amount = GetIniString(EXE_PATH & "\Auto_RTJN_Qty_Amount.ini", "FILE", "lokasi_simpan_file_amount")

        'monitoring qty
        nama_file_template_monitoring_qty = GetIniString(EXE_PATH & "\Auto_RTJN_Qty_Amount.ini", "MONITORING_QTY", "nama_file_template_monitoring_qty")
        nama_file_lampiran_email_monitoring_qty = GetIniString(EXE_PATH & "\Auto_RTJN_Qty_Amount.ini", "MONITORING_QTY", "nama_file_lampiran_email_monitoring_qty")
        lokasi_simpan_file_monitoring_qty = GetIniString(EXE_PATH & "\Auto_RTJN_Qty_Amount.ini", "MONITORING_QTY", "lokasi_simpan_file_monitoring_qty")
        subject_email_monitoring_qty = GetIniString(EXE_PATH & "\Auto_RTJN_Qty_Amount.ini", "MONITORING_QTY", "subject_email_monitoring_qty")
        email_monitoring_mail_sender_qty = GetIniString(EXE_PATH & "\Auto_RTJN_Qty_Amount.ini", "MONITORING_QTY", "email_monitoring_mail_sender_qty")

        'monitoring amount
        nama_file_template_monitoring_amount = GetIniString(EXE_PATH & "\Auto_RTJN_Qty_Amount.ini", "MONITORING_AMOUNT", "nama_file_template_monitoring_amount")
        nama_file_lampiran_email_monitoring_amount = GetIniString(EXE_PATH & "\Auto_RTJN_Qty_Amount.ini", "MONITORING_AMOUNT", "nama_file_lampiran_email_monitoring_amount")
        lokasi_simpan_file_monitoring_amount = GetIniString(EXE_PATH & "\Auto_RTJN_Qty_Amount.ini", "MONITORING_AMOUNT", "lokasi_simpan_file_monitoring_amount")
        subject_email_monitoring_amount = GetIniString(EXE_PATH & "\Auto_RTJN_Qty_Amount.ini", "MONITORING_AMOUNT", "subject_email_monitoring_amount")
        email_monitoring_mail_sender_amount = GetIniString(EXE_PATH & "\Auto_RTJN_Qty_Amount.ini", "MONITORING_AMOUNT", "email_monitoring_mail_sender_amount")


    End Sub

    Public Shared Function OpenConn(ByVal IPServer As String) As Boolean
        Cn = New SqlConnection(IPServer)
        Cn.Open()

        If Cn.State <> ConnectionState.Open Then
            Return False
        Else
            Return True
        End If
    End Function

    Public Shared Sub CloseConn()
        If Not IsNothing(Cn) Then
            Cn.Close()
            Cn = Nothing
        End If
    End Sub

    Public Shared Function ExecuteQuery(ByVal Query As String, ByVal IPServer As String) As DataTable
        If Not OpenConn(IPServer) Then
            MsgBox("Koneksi Gagal..!!", MsgBoxStyle.Critical, "Access Failed")
            Return Nothing
            Exit Function
        End If

        Cmd = New SqlCommand(Query, Cn)
        Da = New SqlDataAdapter
        Da.SelectCommand = Cmd

        Ds = New Data.DataSet
        Cmd.CommandTimeout = 1000
        Da.Fill(Ds)
        Dt = Ds.Tables(0)

        Ds = Nothing
        Da = Nothing
        Cmd = Nothing

        CloseConn()

        Return Dt

        Dt = Nothing
    End Function

    Public Shared Sub ExecuteNonQuery(ByVal Query As String, ByVal IPServer As String)
        If Not OpenConn(IPServer) Then
            MsgBox("Koneksi Gagal..!!", MsgBoxStyle.Critical, "Access Failed..!!")
            Exit Sub
        End If

        Cmd = New SqlCommand
        Cmd.Connection = Cn
        Cmd.CommandTimeout = 600
        Cmd.CommandType = CommandType.Text
        Cmd.CommandText = Query
        Cmd.ExecuteNonQuery()
        Cmd = Nothing
        CloseConn()
    End Sub

    Public Shared Sub create_log_error(ByVal currentDate As String,
                                       ByVal jenis_laporan As String,
                                       ByVal pesan_error As String)
        Dim PathFile As String = ClsConfig.nama_folder_log_error
        If Not System.IO.Directory.Exists(PathFile) Then
            System.IO.Directory.CreateDirectory(PathFile)
        End If

        Dim nama_file_txt_log_error_n_path As String
        nama_file_txt_log_error_n_path = PathFile & "\" & ClsConfig.nama_file_txt_log_error & ".txt"

        If Not File.Exists(nama_file_txt_log_error_n_path) Then
            Using writer As New StreamWriter(nama_file_txt_log_error_n_path, True)
                writer.Write(pesan_error)
            End Using
        Else
            File.AppendAllText(nama_file_txt_log_error_n_path, Environment.NewLine + pesan_error)
        End If

        'Get Date & Time 
        'Dim current_date As DateTime = Now
        'Dim current_HHmm As Integer = Int32.Parse(current_date.ToString("HHmm"))
        'Dim start_date As DateTime = DateSerial(Year(current_date), Month(current_date), 1)
        'Dim end_date As DateTime = ClsGeneral.get_last_date(current_date)


        'Update file maintenance
        'Dim ClsGnrl As New ClsGeneral
        'ClsGnrl.monitoring_email(currentDate, current_HHmm, jenis_laporan, pesan_error)

        'Send Maintenance Report
        'Dim ClsAutRep As New ClsAutoReport
        'ClsAutRep.AutoReportMonitoring(start_date, current_date, "qty") 'Send Email Maintenance Qty
        'ClsAutRep.AutoReportMonitoring(start_date, current_date, "qty_amount") 'Send Email Maintenance Qty_Amount
        'ClsAutRep.AutoReportMonitoring(start_date, current_date, "all") 'Send Email Maintenance Qty


    End Sub

End Class
