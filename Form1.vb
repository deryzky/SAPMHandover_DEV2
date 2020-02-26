Imports System.Configuration
Imports System.IO
Imports System.IO.Compression
Imports Microsoft.Office.Interop
Imports Teradata.Client.Provider
Imports SAPMHandover_DEV2.Module1
Imports System.Messaging

Public Class Form1


    Function GetAppKey(ByVal myKey As String) As String
        Try
            Dim appSettings = ConfigurationManager.AppSettings
            GetAppKey = appSettings(myKey)
        Catch e As ConfigurationErrorsException
            Console.WriteLine("Error reading app settings")
            GetAppKey = ""
        End Try
    End Function
    Sub clear()
        lblHitungDataConvert.Text = ""
        lblWaktuMulai.Text = ""
        lblWaktuMulaiConvert.Text = ""
        lblWaktuSelesaiConvert.Text = ""
        lblWaktuSelesai.Text = ""
        lblWilayah.Text = ""
        lblRecord.Text = ""
        lblCabang.Text = ""
        txtLokasiFileWinrar.Text = "c:\program files\winrar\winrar.exe"
    End Sub
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Call clear()
        For i As Integer = 1 To 12
            cmbBulan.Items.Add(MonthName(i, False))
        Next

        For i = 2014 To 2024
            cmbTahun.Items.Add(i)
        Next i
    End Sub
    Private Function DataIsOK() As Boolean
        DataIsOK = True

        If Trim(txtNamaTabel.Text) = "" Then
            DataIsOK = False
            MsgBox("Nama Database dan Nama Tabel belum diisi", vbExclamation, Me.Text)
            Exit Function
        End If

        If Trim(cmbBulan.Text) = "" Then
            DataIsOK = False
            MsgBox("Bulan Periode belum diisi", vbExclamation, Me.Text)
            Exit Function
        End If

        If Trim(cmbTahun.Text) = "" Then
            DataIsOK = False
            MsgBox("Tahun Periode belum diisi", vbExclamation, Me.Text)
            Exit Function
        End If

        If Trim(txtLokasiFolderKB.Text) = "" Then
            DataIsOK = False
            MsgBox("Lokasi Folder KB belum diisi", vbExclamation, Me.Text)
            Exit Function
        End If
    End Function
    Private Sub cmdProses_Click(sender As Object, e As EventArgs) Handles cmdProses.Click
        If DataIsOK() Then
            If MsgBox("Proses Data " & (Trim(cmbBulan.Text & " " & cmbTahun.Text & " ?")),
        vbQuestion + vbYesNo, "Confirmation") = vbYes Then
                lblWaktuMulai.Text = Now
                lblWaktuMulaiConvert.Text = Now
                headerdat()
                createDirWilayah()
                total()
                lblWaktuSelesaiConvert.Text = Now
                lblWaktuSelesai.Text = Now
                MsgBox("Selesai")
            End If
        End If
    End Sub
    Private Sub cmdClose_Click(sender As Object, e As EventArgs) Handles cmdClose.Click
        Dispose()
    End Sub

    Private Sub connection(ByRef conn As TdConnection)
        Dim connectionString As String = GetAppKey("CONN_STR")
        'Dim conn As TdConnection
        Try

            conn = New TdConnection(connectionString)
            conn.Open()
            'Console.WriteLine(myText & "Connection opened, Process Dataset" & vbCrLf)

        Catch ex As TdException
            'Console.WriteLine(myText & "Error: " & ex.ToString & vbCrLf)
        End Try
    End Sub
    Public Sub createDirWilayah()
        Dim conn As TdConnection
        connection(conn)
        'Set up connection string
        Me.Cursor = Cursors.WaitCursor
        Dim wilayah As String
        Dim myText As String = "Connecting to database using teradata" & vbCrLf & vbCrLf
        Dim connectionString As String = GetAppKey("CONN_STR")
        Dim query As String = GetAppKey("QUERY")
        'Dim query2 As String = GetAppKey("QUERY", "Where Region_Kelolaan = '" & mystr & "'")

        Try


            Console.WriteLine(myText & "Connection opened, Process Dataset" & vbCrLf)

            Dim tout As Integer = CInt(GetAppKey("TIMEOUT"))
            Dim command = New TdCommand(query, conn)
            command.CommandTimeout = tout
            Dim reader As TdDataReader = command.ExecuteReader
            Dim i As Integer
            Dim strDlm As String = GetAppKey("DELIMITER")

            Dim mystrHeader As String = ""
            Dim iHeader As Integer = 0


            While reader.Read()

                wilayah = ""

                For i = 0 To reader.FieldCount - 1

                    If GetAppKey("HEADER") = "Y" And iHeader = 0 Then
                        If i < reader.FieldCount - 1 Then
                            mystrHeader = mystrHeader & reader.GetName(i) & strDlm
                        Else
                            mystrHeader = mystrHeader & reader.GetName(i)
                        End If
                    End If
                    If i < reader.FieldCount - 1 Then
                        wilayah = wilayah & reader.GetValue(i) & strDlm
                    Else
                        wilayah = wilayah & reader.GetValue(i)
                    End If
                Next

                If GetAppKey("HEADER") = "Y" And iHeader = 0 Then
                    iHeader = 1
                End If

                If Dir(Trim(txtLokasiFolderKB.Text) & "FileExcel", vbDirectory) = "" Then
                    MkDir(Trim(txtLokasiFolderKB.Text) & "FileExcel")
                End If
                If Dir(Trim(txtLokasiFolderKB.Text) & "FileExcel\6" & String.Format("{0:00}", CInt(wilayah)), vbDirectory) = "" Then
                    MkDir(Trim(txtLokasiFolderKB.Text) & "FileExcel\6" & String.Format("{0:00}", CInt(wilayah)))
                End If
                If Dir(Trim(txtLokasiFolderKB.Text) & "6" & String.Format("{0:00}", CInt(wilayah)), vbDirectory) = "" Then
                    MkDir(Trim(txtLokasiFolderKB.Text) & "6" & String.Format("{0:00}", CInt(wilayah)))
                End If

                createDirCabang(wilayah)
            End While

            reader.Close()
            conn.Close()
            'Console.WriteLine(myText & "Connection closed." & vbCrLf)

        Catch ex As TdException
            Console.WriteLine(myText & "Error: " & ex.ToString & vbCrLf)
        End Try
        Me.Cursor = Cursors.Default
    End Sub
    Public Sub createExcel(ByRef wilayah As String, ByRef cabang As String, ByRef nmCbng As String)
        Dim myText As String = "Connecting to database using teradata" & vbCrLf & vbCrLf
        Dim strConn As String = ""
        Dim strConnAttr As String = ""
        Dim strConnVal As String = ""
        Dim connectionString As String = GetAppKey("CONN_STR")
        Dim conn As TdConnection
        Dim query As String

        Dim bulan As String
        bulan = Convert_Date_Str2Int(cmbBulan.Text)

        Dim sNmaFileTxt2 As String
        sNmaFileTxt2 = Trim(txtLokasiFolderKB.Text) & "6" & String.Format("{0:00}", CInt(wilayah)) & "\" &
            String.Format("{0:000}", CInt(cabang)) & "\Cabang_" & wilayah & "_" & cabang & "_" & Trim(nmCbng) &
            "_" & cmbTahun.Text & Trim(bulan) & ".txt"

        Dim periode As String
        periode = cmbTahun.Text & "-" & bulan & "-" & "31"

        Dim args_query(3) As String
        args_query(0) = periode
        args_query(1) = wilayah
        args_query(2) = cabang
        query = QueryBuilder(args_query, (GetAppKey("QUERYTXTCABANG")))

        Try

            conn = New TdConnection(connectionString)
            conn.Open()
            Console.WriteLine(myText & "Connection opened, Process Dataset" & vbCrLf)
            'Console.WriteLine(myQuery)

            Dim xlWorkSheet As Excel.Worksheet
            Dim xlApp As Excel.Application
            Dim xlWorkBook As Excel.Workbook
            Dim misValue As Object = System.Reflection.Missing.Value

            xlApp = New Excel.Application
            'xlApp = New Excel.ApplicationClass
            xlWorkBook = xlApp.Workbooks.Add(misValue)
            xlWorkSheet = xlWorkBook.Sheets("sheet1")

            Dim tout As Integer = CInt(GetAppKey("TIMEOUT"))
            Dim cmd = New TdCommand(query, conn)
            cmd.CommandTimeout = tout
            Dim read As TdDataReader = cmd.ExecuteReader
            Dim i As Integer
            Dim strDlm As String = GetAppKey("DELIMITER")
            Dim mystr As String = ""
            Dim mystrHeader As String = ""
            Dim iHeader As Integer = 0
            Dim row As Integer = 0
            'Console.WriteLine("Create Textfile: " & MyFileName & vbCrLf)
            'Using sw As StreamWriter = New StreamWriter(MyFileName)

            While read.Read()

                mystr = ""
                'reader.FieldCount
                For i = 0 To read.FieldCount - 1
                    'Console.WriteLine("{0} = {1}", reader.GetName(i), reader.GetValue(i))
                    If GetAppKey("HEADER") = "Y" And iHeader = 0 Then
                        If i < read.FieldCount - 1 Then
                            xlWorkSheet.Cells(1, i + 1) = read.GetName(i).ToString
                        End If
                    End If
                    xlWorkSheet.Cells(row + 1, i + 1) = read.GetValue(i).ToString
                Next



                row = row + 1

            End While
            My.Computer.FileSystem.DeleteFile(sNmaFileTxt2)
            Dim dirXl As String = Trim(txtLokasiFolderKB.Text) & "FileExcel\6" &
            String.Format("{0:00}", CInt(wilayah)) & "\" & String.Format("{0:000}", CInt(cabang))
            Dim namaFileXl As String = "Cabang_" & wilayah & "_" & cabang & "_" & Trim(nmCbng) & "_" &
            cmbTahun.Text & Trim(bulan) & ".xlsx"

            xlWorkSheet.Cells.EntireColumn.AutoFit()
            xlWorkSheet.SaveAs(dirXl & "\" & namaFileXl)
            xlWorkBook.Close()
            xlApp.Quit()

            releaseObject(xlApp)
            releaseObject(xlWorkBook)
            releaseObject(xlWorkSheet)
            read.Close()
            conn.Close()
        Catch ex As TdException
            Console.WriteLine(myText & "Error:  " & ex.ToString & vbCrLf)
        End Try
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
    Public Sub createDirCabang(ByRef wilayah As String)
        Dim conn As TdConnection
        connection(conn)

        Dim cabang As String
        Dim myText As String = "Connecting to database using teradata" & vbCrLf & vbCrLf

        Dim connectionString As String = GetAppKey("CONN_STR")
        Dim query As String
        Dim args_query(1) As String
        args_query(0) = wilayah
        query = QueryBuilder(args_query, (GetAppKey("QUERY2")))

        Try


            Console.WriteLine(myText & "Connection opened, Process Dataset" & vbCrLf)
            'Console.WriteLine(myQuery)

            Dim tout As Integer = CInt(GetAppKey("TIMEOUT"))
            Dim command = New TdCommand(query, conn)
            command.CommandTimeout = tout
            Dim reader As TdDataReader = command.ExecuteReader
            Dim i As Integer
            Dim strDlm As String = GetAppKey("DELIMITER")
            Dim mystrHeader As String = ""
            Dim iHeader As Integer = 0
            Dim a As Integer

            a = 1
            While reader.Read()
                cabang = ""
                For i = 0 To reader.FieldCount - 1
                    'Console.WriteLine("{0} = {1}", reader.GetName(i), reader.GetValue(i))
                    If GetAppKey("HEADER") = "Y" And iHeader = 0 Then
                        If i < reader.FieldCount - 1 Then
                            mystrHeader = mystrHeader & reader.GetName(i) & strDlm
                        Else
                            mystrHeader = mystrHeader & reader.GetName(i)
                        End If
                    End If
                    If i < reader.FieldCount - 1 Then
                        cabang = cabang & reader.GetValue(i) & strDlm
                    Else
                        cabang = cabang & reader.GetValue(i)
                    End If
                Next

                If GetAppKey("HEADER") = "Y" And iHeader = 0 Then
                    iHeader = 1
                End If

                If Dir(Trim(txtLokasiFolderKB.Text) & "FileExcel\6" & String.Format("{0:00}", CInt(wilayah)) &
                    "\" & String.Format("{0:000}", CInt(cabang)), vbDirectory) = "" Then
                    MkDir(Trim(txtLokasiFolderKB.Text) & "FileExcel\6" & String.Format("{0:00}", CInt(wilayah)) &
                    "\" & String.Format("{0:000}", CInt(cabang)))
                End If
                If Dir(Trim(txtLokasiFolderKB.Text) & "6" & String.Format("{0:00}", CInt(wilayah)) & "\" &
                    String.Format("{0:000}", CInt(cabang)), vbDirectory) = "" Then
                    MkDir(Trim(txtLokasiFolderKB.Text) & "6" & String.Format("{0:00}", CInt(wilayah)) & "\" &
                    String.Format("{0:000}", CInt(cabang)))
                End If

                Dim sNamaFileZip As String
                Dim sPassword As String
                Dim bulan As String
                bulan = Convert_Date_Str2Int(cmbBulan.Text)
                sPassword = "Pwd" & wilayah & cabang & cmbTahun.Text & (Trim(bulan))
                sNamaFileZip = Trim(txtLokasiFolderKB.Text) & "6" & String.Format("{0:00}", CInt(wilayah)) &
                "\Wilayah_" & wilayah & "_" & cmbTahun.Text & Trim(bulan) & "_" & a & ".zip"
                a = a + 1

                'Dim lokasiWinrar As String = txtLokasiFileWinrar.Text
                'Dim sNamaFileAsli As String = "C:\DataSAPM\test.txt"

                getNamaCabangKelolaan(wilayah, cabang)

                'txttext(wilayah, cabang)
                'gpCompressFileToZip(lokasiWinrar, sNamaFileAsli, sPassword, sNamaFileZip)

            End While

            '    sw.Close()
            'End Using
            reader.Close()
            conn.Close()
            Console.WriteLine(myText & "Connection closed." & vbCrLf)

        Catch ex As TdException
            Console.WriteLine(myText & "Error: " & ex.ToString & vbCrLf)
        End Try

    End Sub

    Function QueryBuilder(ByVal args() As String, ByVal MyQuery As String) As String
        Try
            Dim str As String = MyQuery
            If args.Count >= 1 Then
                For i As Integer = 0 To args.Length - 1
                    'Console.WriteLine("@ARGS" & (i + 1).ToString)
                    str = str.Replace("@ARGS" & (i + 1).ToString, args(i))
                    'Console.WriteLine(str)
                Next
                QueryBuilder = str
            Else
                QueryBuilder = MyQuery
            End If
        Catch ex As Exception
            Console.WriteLine("QueryBuilder error: " & ex.ToString)
            QueryBuilder = MyQuery
        End Try
    End Function

    Private Sub getNamaCabangKelolaan(ByRef wilayah As String, ByRef cabang As String)
        Dim conn As TdConnection
        connection(conn)

        Dim myText As String = "Connecting to database using teradata" & vbCrLf & vbCrLf
        Dim strConn As String = ""
        Dim strConnAttr As String = ""
        Dim strConnVal As String = ""
        Dim connectionString As String = GetAppKey("CONN_STR")
        Dim query As String


        Dim args_query(1) As String
        args_query(0) = wilayah
        args_query(1) = cabang
        query = QueryBuilder(args_query, (GetAppKey("QUERY3")))

        Try

            Console.WriteLine(myText & "Connection opened, Process Dataset" & vbCrLf)
            'Console.WriteLine(myQuery)

            Dim tout As Integer = CInt(GetAppKey("TIMEOUT"))
            Dim command = New TdCommand(query, conn)
            command.CommandTimeout = tout
            Dim reader As TdDataReader = command.ExecuteReader
            Dim i As Integer
            Dim strDlm As String = GetAppKey("DELIMITER")
            Dim nmCbng As String = ""
            Dim mystrHeader As String = ""
            Dim iHeader As Integer = 0
            'Console.WriteLine("Create Textfile: " & MyFileName & vbCrLf)

            While reader.Read()
                nmCbng = ""
                'reader.FieldCount
                For i = 0 To reader.FieldCount - 1
                    'Console.WriteLine("{0} = {1}", reader.GetName(i), reader.GetValue(i))
                    If GetAppKey("HEADER") = "Y" And iHeader = 0 Then
                        If i < reader.FieldCount - 1 Then
                            mystrHeader = mystrHeader & reader.GetName(i) & strDlm
                        Else
                            mystrHeader = mystrHeader & reader.GetName(i)
                        End If
                        'ts.WriteLine(mystrHeader)
                    End If
                    If i < reader.FieldCount - 1 Then
                        nmCbng = nmCbng & reader.GetValue(i) & strDlm
                    Else
                        nmCbng = nmCbng & reader.GetValue(i)
                    End If
                Next
                nmCbng = Replace(nmCbng, " ", "")
                createTxt(wilayah, cabang, nmCbng)
                'createExcel(wilayah, cabang, nmCbng)
            End While
            'reader.Close()
            'conn.Close()
            Console.WriteLine(myText & "Connection closed." & vbCrLf)
        Catch ex As TdException
            Console.WriteLine(myText & "Error: " & ex.ToString & vbCrLf)
        End Try

    End Sub

    Private Sub IsiBNISegmentasiDAT(ByRef mystr As String, ByRef mystr2 As String, ByRef sNamaFileZip As String,
                                    Optional ByRef sPassword As String = "")
        Dim fso As New Scripting.FileSystemObject
        Dim ts As Scripting.TextStream
        Dim filenya As Scripting.File

        filenya = fso.GetFile(Trim(txtLokasiFolderKB.Text) & "bnisegmentasi.dat")
        ts = filenya.OpenAsTextStream(Scripting.IOMode.ForAppending)

        'ts.WriteLine "6" & Format(sKodeWilayah, "00") & "|" & Format(sKodeCabang, "000") & "|" & sNamaFileZip & " & _ & a &|" & sPassword
        ts.WriteLine("6" & String.Format("{0:00}", CInt(mystr)) & "|" & String.Format("{0:0}", CInt(mystr2)) &
        "|" & sNamaFileZip & "|" & sPassword)
        'ts.WriteLine "Format(sKodeWilayah, "00") & "|" & sKodeCabang & "|" & sNamaFileZip & "|" & sPassword 'FORMAT LAMA

    End Sub

    Public Sub gpCompressFileToZip(ByRef sLokasiWinrar As String, ByRef sNmaFileTxt As String, ByRef sPassword As String,
                                   ByRef sNamaFileZip As String)
        Shell(sLokasiWinrar & " a -p" & sPassword & " " & sNamaFileZip & " " & sNmaFileTxt)
    End Sub

    Private Sub createTxt(ByRef wilayah As String, ByRef cabang As String, ByRef nmCbng As String)
        Dim conn As TdConnection
        connection(conn)

        Dim myText As String = "Connecting to database using teradata" & vbCrLf & vbCrLf

        Dim connectionString As String = GetAppKey("CONN_STR")
        'Dim query As String = GetAppKey("QUERYTXTCABANG")
        Dim query As String


        Dim bulan As String
        bulan = Convert_Date_Str2Int(cmbBulan.Text)
        Dim periode As String
        periode = cmbTahun.Text & "-" & bulan & "-" & "31"

        Dim args_query(3) As String
        args_query(0) = periode
        args_query(1) = wilayah
        args_query(2) = cabang
        query = QueryBuilder(args_query, (GetAppKey("QUERYTXTCABANG")))

        Try

            Console.WriteLine(myText & "Connection opened, Process Dataset" & vbCrLf)
            'Console.WriteLine(myQuery)

            Dim tout As Integer = CInt(GetAppKey("TIMEOUT"))
            Dim command = New TdCommand(query, conn)
            command.CommandTimeout = tout
            Dim reader As TdDataReader = command.ExecuteReader
            Dim i As Integer
            Dim strDlm As String = GetAppKey("DELIMITER")
            Dim mystr As String = ""
            Dim mystrHeader As String = ""
            Dim iHeader As Integer = 0
            'Console.WriteLine("Create Textfile: " & MyFileName & vbCrLf)
            'Using sw As StreamWriter = New StreamWriter(MyFileName)

            Dim sNmaFileTxt2 As String
            sNmaFileTxt2 = Trim(txtLokasiFolderKB.Text) & "6" & String.Format("{0:00}", CInt(wilayah)) & "\" &
            String.Format("{0:000}", CInt(cabang)) & "\Cabang_" & wilayah & "_" & cabang & "_" & Trim(nmCbng) &
            "_" & cmbTahun.Text & Trim(bulan) & ".txt"

            Dim sNmaFileTxt As String

            ChDir(Trim(txtLokasiFolderKB.Text) & "6" & String.Format("{0:00}", CInt(wilayah)) & "\" &
            String.Format("{0:000}", CInt(cabang)) & "\")
            sNmaFileTxt = "Cabang_" & wilayah & "_" & cabang & "_" & Trim(nmCbng) & "_" & cmbTahun.Text &
            Trim(bulan) & ".txt"
            Dim ts As System.IO.StreamWriter
            ts = My.Computer.FileSystem.OpenTextFileWriter(sNmaFileTxt, True)


            While reader.Read()

                mystr = ""

                For i = 0 To reader.FieldCount - 1
                    'Console.WriteLine("{0} = {1}", reader.GetName(i), reader.GetValue(i))
                    If GetAppKey("HEADER") = "Y" And iHeader = 0 Then
                        If i < reader.FieldCount - 1 Then
                            mystrHeader = mystrHeader & reader.GetName(i) & strDlm
                        Else
                            mystrHeader = mystrHeader & reader.GetName(i)
                        End If
                    End If
                    If i < reader.FieldCount - 1 Then
                        mystr = mystr & reader.GetValue(i) & strDlm
                    Else
                        mystr = mystr & reader.GetValue(i)
                    End If
                Next

                If GetAppKey("HEADER") = "Y" And iHeader = 0 Then
                    iHeader = 1
                    ts.WriteLine(mystrHeader)
                End If

                ts.WriteLine(mystr)

            End While
            ts.Close()
            Dim sNamaFileZip As String
            sNamaFileZip = Trim(txtLokasiFolderKB.Text) & "6" & String.Format("{0:00}", CInt(wilayah)) & "\" &
                           String.Format("{0:000}", CInt(cabang)) & "\Cabang_" & wilayah & "_" & cabang & "_" &
                           Trim(nmCbng) & "_" & cmbTahun.Text & Trim(bulan) & ".zip"
            Dim sPassword As String
            sPassword = "Pwd" & wilayah & cabang & cmbTahun.Text & (Trim(bulan))
            Dim lokasiWinrar As String = txtLokasiFileWinrar.Text

            If Dir((Trim(txtLokasiFolderKB.Text) & "6" & String.Format("{0:00}", CInt(wilayah)) & "\" &
                   String.Format("{0:000}", CInt(cabang)) & "\Cabang_" & wilayah & "_" & cabang & "_" & Trim(nmCbng) &
                   "_" & cmbTahun.Text & Trim(bulan) & ".zip"), vbDirectory) = "" Then
                gpCompressFileToZip(lokasiWinrar, sNmaFileTxt, sPassword, sNamaFileZip)
            End If

            Dim sNamaFileZip2 As String
            ChDir(Trim(txtLokasiFolderKB.Text) & "6" & String.Format("{0:00}", CInt(wilayah)) & "\" &
                  String.Format("{0:000}", CInt(cabang)) & "\")
            sNamaFileZip2 = "Cabang_" & wilayah & "_" & cabang & "_" & Trim(nmCbng) & "_" & cmbTahun.Text &
                Trim(bulan) & ".zip"

            CreateBNISegmentasiDAT(wilayah, cabang, sNamaFileZip2, sPassword)

            createExcel(wilayah, cabang, nmCbng)

            'My.Computer.FileSystem.DeleteFile(sNmaFileTxt2)


            ' sw.Close()
            '  End Using
            reader.Close()
            conn.Close()
            Console.WriteLine(myText & "Connection closed." & vbCrLf)

        Catch ex As TdException
            Console.WriteLine(myText & "Error: " & ex.ToString & vbCrLf)
        End Try

    End Sub

    Private Sub CreateBNISegmentasiDAT(ByRef wilayah As String, ByRef cabang As String, ByRef sNamaFileZip As String,
                                       ByRef sPassword As String)
        Dim fso As New Object
        Dim ts As System.IO.StreamWriter
        ts = My.Computer.FileSystem.OpenTextFileWriter(Trim(txtLokasiFolderKB.Text) & "bnisegmentasi.dat", True)
        'ts.WriteLine("Region_Code|Branch_Code|File_Name|Password_File")
        ts.WriteLine(wilayah & "|" & cabang & "|" & sNamaFileZip & "|" & sPassword)
        ts.Close()
    End Sub

    Private Sub headerdat()
        Dim ts As System.IO.StreamWriter
        ts = My.Computer.FileSystem.OpenTextFileWriter(Trim(txtLokasiFolderKB.Text) & "bnisegmentasi.dat", True)
        ts.WriteLine("Region_Code|Branch_Code|File_Name|Password_File")
        ts.Close()
    End Sub


    Private Sub total()
        Dim wilayah As String
        Dim cabang As String
        Dim totalRec As String
        Dim myText As String = "Connecting to database using teradata" & vbCrLf & vbCrLf
        Dim connectionString As String = GetAppKey("CONN_STR")
        Dim queryWil As String = GetAppKey("QUERYTOTALWILAYAH")
        Dim queryCab As String = GetAppKey("QUERYTOTALCABANG")
        Dim queryRec As String = GetAppKey("QUERYTOTALRECORD")
        Dim conn As TdConnection
        Try

            conn = New TdConnection(connectionString)
            conn.Open()
            Console.WriteLine(myText & "Connection opened, Process Dataset" & vbCrLf)

            Dim tout As Integer = CInt(GetAppKey("TIMEOUT"))
            Dim commandWil = New TdCommand(queryWil, conn)
            commandWil.CommandTimeout = tout
            Dim readerWil As TdDataReader = commandWil.ExecuteReader
            Dim i As Integer
            Dim strDlm As String = GetAppKey("DELIMITER")

            Dim commandCab = New TdCommand(queryCab, conn)
            commandCab.CommandTimeout = tout
            Dim readerCab As TdDataReader = commandCab.ExecuteReader
            Dim j As Integer

            Dim commandRec = New TdCommand(queryRec, conn)
            commandRec.CommandTimeout = tout
            Dim readerRec As TdDataReader = commandRec.ExecuteReader
            Dim k As Integer

            Dim mystrHeader As String = ""
            Dim iHeader As Integer = 0


            While readerWil.Read()

                wilayah = ""

                For i = 0 To readerWil.FieldCount - 1
                    wilayah = readerWil.GetValue(i)
                Next

            End While

            While readerCab.Read()

                cabang = ""

                For j = 0 To readerCab.FieldCount - 1
                    cabang = readerCab.GetValue(j)
                Next

            End While

            While readerRec.Read()

                totalRec = ""

                For k = 0 To readerRec.FieldCount - 1
                    totalRec = readerRec.GetValue(k)
                Next

            End While

            lblRecord.Text = totalRec
            lblWilayah.Text = wilayah
            lblCabang.Text = cabang
            readerWil.Close()
            conn.Close()
        Catch ex As TdException
            Console.WriteLine(myText & "Error: " & ex.ToString & vbCrLf)
        End Try
    End Sub



End Class
