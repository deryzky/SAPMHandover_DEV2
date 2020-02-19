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
        'If MsgBox("Proses Data " & Trim(GetHashCode(cmbBulan.Text, 1, "||")) & " " & cmbTahun.Text & " ?", vbQuestion + vbYesNo, "Confirmation") = vbYes Then
        'If MsgBox("Proses Data " & (Trim(cmbBulan.Text & " " & cmbTahun.Text & " ?")),
        'vbQuestion + vbYesNo, "Confirmation") = vbYes Then

        '        If DataIsOK() Then
        '            lblWaktuMulai.Text = Now
        '            'Me.MousePosition = 
        '            'Application.DoEvents()
        '            'CreateBNISegmentasiDAT()
        Dim mystr As String
        Export_Only_Wilayah_UpdateDita(mystr)
        'Export_Only_Cabang()

        '            'MsgBox(Convert_Date_Str2Int(cmbBulan.Text))
        '            lblWaktuSelesai.Text = Now
        '            'Me.MousePointer = vbNormal
        '            MsgBox("Selesai")
        '        End If
        '    End If
    End Sub
    Private Sub cmdClose_Click(sender As Object, e As EventArgs) Handles cmdClose.Click
        End
    End Sub
    Private Sub CreateBNISegmentasiDAT()
        Dim fso As New Object
        Dim ts As System.IO.StreamWriter
        ts = My.Computer.FileSystem.OpenTextFileWriter(Trim(txtLokasiFolderKB.Text) & "bnisegmentasi.dat", True)
        ts.WriteLine("Region_Code|Branch_Code|File_Name|Password_File")
        ts.Close()
    End Sub
    Public Sub Export_Only_Wilayah_UpdateDita(ByRef mystr As String)
        'Set up connection string
        Dim myText As String = "Connecting to database using teradata" & vbCrLf & vbCrLf
        Dim strConn As String = ""
        Dim strConnAttr As String = ""
        Dim strConnVal As String = ""
        Dim connectionString As String = GetAppKey("CONN_STR")
        Dim query As String = GetAppKey("QUERY")
        'Dim query2 As String = GetAppKey("QUERY", "Where Region_Kelolaan = '" & mystr & "'")

        Dim conn As TdConnection
        Dim sSQL2 As String


        Try

            conn = New TdConnection(connectionString)
            conn.Open()
            Console.WriteLine(myText & "Connection opened, Process Dataset" & vbCrLf)
            'Console.WriteLine(myQuery)

            Dim tout As Integer = CInt(GetAppKey("TIMEOUT"))
            Dim command = New TdCommand(query, conn)
            command.CommandTimeout = tout
            Dim reader As TdDataReader = command.ExecuteReader
            Dim i As Integer
            Dim strDlm As String = GetAppKey("DELIMITER")
            'Dim mystr As String = ""
            Dim mystrHeader As String = ""
            Dim iHeader As Integer = 0
            'Console.WriteLine("Create Textfile: " & MyFileName & vbCrLf)
            'Using sw As StreamWriter = New StreamWriter(MyFileName)

            While reader.Read()

                mystr = ""
                'reader.FieldCount
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
                    'sw.WriteLine(mystrHeader)
                End If
                lblWilayah.Text = mystr
                'mystr = IsiField(reader("Region_Kelolaan"))
                'MsgBox(mystr)

                'sw.WriteLine(mystr)

                sSQL2 = "Select * from " & txtNamaTabel.Text & " " &
                    "Where Region_Kelolaan = '" & mystr & "'"

                If Dir(Trim(txtLokasiFolderKB.Text) & "FileExcel", vbDirectory) = "" Then
                    MkDir(Trim(txtLokasiFolderKB.Text) & "FileExcel")
                End If
                If Dir(Trim(txtLokasiFolderKB.Text) & "FileExcel\6" & String.Format("{0:00}", CInt(mystr)), vbDirectory) = "" Then
                    MkDir(Trim(txtLokasiFolderKB.Text) & "FileExcel\6" & String.Format("{0:00}", CInt(mystr)))
                End If
                If Dir(Trim(txtLokasiFolderKB.Text) & "6" & String.Format("{0:00}", CInt(mystr)), vbDirectory) = "" Then
                    MkDir(Trim(txtLokasiFolderKB.Text) & "6" & String.Format("{0:00}", CInt(mystr)))
                End If
                Dim sNamaFileZip As String
                Dim sPassword As String
                Export_Only_Cabang(mystr, sNamaFileZip, sPassword)



                Dim bulan As String
                bulan = Convert_Date_Str2Int(cmbBulan.Text)


                'MsgBox(bulan)
                'End
                ' strop = cmbBulan.Text
                'coba.Text = cmbBulan.Text
                'coba.Text = str
                'MsgBox("sini")
                Dim sNamaFile As String
                Dim sNamaFileZip2 As String
                Dim sNamaFileZipFileInfoData As String
                'Dim sPassword As String
                Dim a As Integer
                Dim b As Integer

                'siapkan nama file dan password zip
                sNamaFile = Trim(txtLokasiFolderKB.Text) & "FileExcel\6" & String.Format("{0:00}", CInt(mystr)) & "\Wilayah_" & mystr & "" & ".xlsx"
                'sPassword = "Pwd" & mystr & cmbTahun.Text & Trim(bulan)
                'MsgBox(sPassword)
                'End
                'sNamaFileZip = Trim(txtLokasiFolderKB.Text) & "6" & String.Format("{0:00}", CInt(mystr)) & "\Wilayah_" & mystr & "_" & cmbTahun.Text & Trim(bulan) & ".zip"
                sNamaFileZipFileInfoData = "Wilayah_" & mystr & "_" & cmbTahun.Text & Trim(bulan) & ".zip"

                sNamaFile = Trim(txtLokasiFolderKB.Text) & "FileExcel\6" & String.Format("{0:00}", CInt(mystr)) & "\Wilayah_" & mystr & "" & ".xlsx"
                sNamaFileZip2 = Trim(txtLokasiFolderKB.Text) & "6" & String.Format("{0:00}", CInt(mystr)) & "\Wilayah_" & mystr & "_" & cmbTahun.Text & Trim(bulan) & ".zip"
                'MsgBox(sNamaFileZip)
                'End
                b = 0
                a = 0

                'sNamaFileZip = (Trim(txtLokasiFolderKB.Text) & "6" & String.Format("{0:00}", CInt(mystr)) & "\Wilayah_" & mystr & "_" & cmbTahun.Text & (Trim(bulan)) & "_" & a & ".zip")

                'sNamaFileZipFileInfoData = "Wilayah_" & mystr & "_" & cmbTahun.Text & (Trim(bulan)) & "_" & a & "" & ".zip"
                'MsgBox(sNamaFileZipFileInfoData)
                'gpCompressFileToZip(txtLokasiFileWinrar.Text, sNamaFile, sPassword, sNamaFileZip)
                'IsiBNISegmentasiDAT(String.Format("{0:00}", CInt(mystr)), "0", sNamaFileZipFileInfoData, sPassword)

                'IsiBNISegmentasiDAT(Format(mystr, "00"), "0", sNamaFileZipFileInfoData, sPassword)
                'MsgBox("Selesai")
                'End
                'gpCompressFileToZip txtLokasiFileWinrar, sNamaFile, sPassword, sNamaFileZip
                'IsiBNISegmentasiDAT Format(sKodeWilayah, "00"), "0", sNamaFileZipFileInfoData, sPassword
                'coba.Text = sNamaFileZipFileInfoData

                'Dim oExcel As Object
                'Dim oBook As Object
                'Dim sht As Object




                sNamaFile = Trim(txtLokasiFolderKB.Text) & "FileExcel\6" & Format(mystr, "00") & "\Wilayah_" & mystr & "_" & a & "" & ".xlsx"
                'coba.Text = sNamaFile
                'oBook.SaveAs.sNamaFile
                'oBook.Close
                'oExcel.Quit

            End While

            Dim queryex As String = GetAppKey("QUERYEXCEL")
            Dim cmd = New TdCommand(queryex, conn)
            Dim read As TdDataReader = cmd.ExecuteReader

            Dim xlApp As Excel.Application
            Dim xlWorkBook As Excel.Workbook
            Dim xlWorkSheet As Excel.Worksheet
            Dim misValue As Object = System.Reflection.Missing.Value

            xlApp = New Excel.Application
            'xlApp = New Excel.ApplicationClass
            xlWorkBook = xlApp.Workbooks.Add(misValue)
            xlWorkSheet = xlWorkBook.Sheets("sheet1")

            'Dim command = New TdCommand(query, conn)
            'command.CommandTimeout = tout
            'Dim reader As TdDataReader = command.ExecuteReader
            xlWorkSheet.PageSetup.CenterHeader = "kon"

            'da = New TdDataAdapter(queryex, conn)
            'da.Fill(ds, "PRD_EDW_RPT_MR_VR.SAPM_DETAIL_NASABAH")

            'For i = 0 To ds.Tables(0).Columns.Count - 1
            '    Dim str As String = ds.Tables(0).Columns(i).ColumnName
            '    xlWorkSheet.Cells(1, i + 1) = str
            'Next

            ''For i = 0 To ds.Tables(0).Rows.Count - 1
            ''        For j = 0 To ds.Tables(0).Columns.Count - 1
            ''            xlWorkSheet.Cells(i + 1, j + 1) =
            ''            ds.Tables(0).Rows(i).Item(j)
            ''        Next
            ''    Next
            'xlWorkSheet.SaveAs("c:\SAPM\KB\vbexcel.xlsx")
            'xlWorkBook.Close()
            'xlApp.Quit()

            'releaseObject(xlApp)
            'releaseObject(xlWorkBook)
            'releaseObject(xlWorkSheet)



            'Dim style As String = "g2"
            'Dim test As String
            'test = String.Format("{0:00}", CInt(mystr))
            ''MsgBox(test)
            'sw.Close()
            'End Using

            reader.Close()
            conn.Close()
            'Console.WriteLine(myText & "Connection closed." & vbCrLf)

        Catch ex As TdException
            Console.WriteLine(myText & "Error: " & ex.ToString & vbCrLf)
        End Try
    End Sub
    Public Sub excel()

        Dim myText As String = "Connecting to database using teradata" & vbCrLf & vbCrLf
        Dim strConn As String = ""
        Dim strConnAttr As String = ""
        Dim strConnVal As String = ""
        Dim connectionString As String = GetAppKey("CONN_STR")
        Dim queryex As String = GetAppKey("QUERYEXCEL")
        'Dim query2 As String = GetAppKey("QUERY", "Where Region_Kelolaan = '" & mystr & "'")

        Dim conn As TdConnection
        Dim sSQL2 As String





        'Dim queryex As String = GetAppKey("QUERYEXCEL")
        'Dim cmd = New TdCommand(queryex, conn)
        'Dim read As TdDataReader = cmd.ExecuteReader

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

            xlWorkSheet.PageSetup.CenterHeader = "kon"

            Dim tout As Integer = CInt(GetAppKey("TIMEOUT"))
            Dim cmd = New TdCommand(queryex, conn)
            cmd.CommandTimeout = tout
            Dim read As TdDataReader = cmd.ExecuteReader
            Dim i As Integer
            Dim strDlm As String = GetAppKey("DELIMITER")
            Dim mystr As String = ""
            Dim mystrHeader As String = ""
            Dim iHeader As Integer = 0
            'Console.WriteLine("Create Textfile: " & MyFileName & vbCrLf)
            'Using sw As StreamWriter = New StreamWriter(MyFileName)

            While read.Read()

                mystr = ""
                'reader.FieldCount
                For i = 0 To read.FieldCount - 1
                    'Console.WriteLine("{0} = {1}", reader.GetName(i), reader.GetValue(i))
                    If GetAppKey("HEADER") = "Y" And iHeader = 0 Then
                        If i < read.FieldCount - 1 Then
                            mystrHeader = mystrHeader & read.GetName(i) & strDlm
                        Else
                            mystrHeader = mystrHeader & read.GetName(i)
                        End If
                    End If
                    If i < read.FieldCount - 1 Then
                        mystr = mystr & read.GetValue(i) & strDlm
                    Else
                        mystr = mystr & read.GetValue(i)
                    End If

                Next

                If GetAppKey("HEADER") = "Y" And iHeader = 0 Then
                    iHeader = 1
                    'sw.WriteLine(mystrHeader)
                End If

                xlWorkSheet.Cells(1, i + 1) = mystrHeader
                MsgBox(mystr)


                Dim da As New TdDataAdapter
                Dim ds As New DataSet

                da = New TdDataAdapter(queryex, conn)
                da.Fill(ds, "PRD_EDW_RPT_MR_VR.SAPM_DETAIL_NASABAH")

                'For i = 0 To ds.Tables(0).Columns.Count - 1
                '    Dim str As String = ds.Tables(0).Columns(i).ColumnName
                '    xlWorkSheet.Cells(1, i + 1) = str

                '    For z = 0 To ds.Tables(0).Rows.Count - 1
                '        'For j = 0 To ds.Tables(0).Columns.Count - 1
                '        xlWorkSheet.Cells(z + 2, i + 1) =
                '            ds.Tables(0).Rows(z).Item(i)
                '        'Next
                '    Next
                'Next

                'For i = 0 To ds.Tables(0).Columns.Count - 1
                '    Dim str As String = ds.Tables(0).Columns(i).ColumnName
                '    xlWorkSheet.Cells(1, i + 1) = str

                '    For z = 0 To ds.Tables(0).Rows.Count - 1
                '        'For j = 0 To ds.Tables(0).Columns.Count - 1
                '        xlWorkSheet.Cells(z + 2, i + 1) =
                '            ds.Tables(0).Rows(z).Item(i)
                '        'Next
                '    Next
                'Next

                xlWorkSheet.SaveAs("c:\SAPM\KB\vbexcel.xlsx")
                xlWorkBook.Close()
                xlApp.Quit()

                releaseObject(xlApp)
                releaseObject(xlWorkBook)
                releaseObject(xlWorkSheet)



                'Dim style As String = "g2"
                'Dim test As String
                'test = String.Format("{0:00}", CInt(mystr))
                'MsgBox(test)
                'sw.Close()
                'End Using
                'read.Close()
                conn.Close()
            End While
        Catch ex As TdException
            Console.WriteLine(myText & "Error: " & ex.ToString & vbCrLf)
        End Try
    End Sub
    Private Sub RunSQLReader()
        'Set up connection string
        Dim myText As String = "Connecting to database using teradata" & vbCrLf & vbCrLf
        Dim strConn As String = ""
        Dim strConnAttr As String = ""
        Dim strConnVal As String = ""
        Dim connectionString As String = GetAppKey("CONN_STR")
        Dim conn As TdConnection
        Dim xlWorkSheet As Excel.Worksheet
        Dim xlApp As Excel.Application
        Dim xlWorkBook As Excel.Workbook
        Dim misValue As Object = System.Reflection.Missing.Value

        xlApp = New Excel.Application
        'xlApp = New Excel.ApplicationClass
        xlWorkBook = xlApp.Workbooks.Add(misValue)
        xlWorkSheet = xlWorkBook.Sheets("sheet1")

        xlWorkSheet.PageSetup.CenterHeader = "kon"

        'For i = 0 To ds.Tables(0).Columns.Count - 1
        '    Dim str As String = ds.Tables(0).Columns(i).ColumnName
        '    xlWorkSheet.Cells(1, i + 1) = str

        '    For z = 0 To ds.Tables(0).Rows.Count - 1
        '        'For j = 0 To ds.Tables(0).Columns.Count - 1
        '        xlWorkSheet.Cells(z + 2, i + 1) =
        '            ds.Tables(0).Rows(z).Item(i)
        '        'Next
        '    Next
        'Next

        Try

            conn = New TdConnection(connectionString)
            conn.Open()
            Console.WriteLine(myText & "Connection opened, Process Dataset" & vbCrLf)
            'Console.WriteLine(myQuery)
            Dim queryex As String = GetAppKey("QUERYCOBA")
            Dim tout As Integer = CInt(GetAppKey("TIMEOUT"))
            Dim command = New TdCommand(queryex, conn)
            command.CommandTimeout = tout
            Dim reader As TdDataReader = command.ExecuteReader
            Dim i As Integer
            Dim strDlm As String = GetAppKey("DELIMITER")
            Dim mystr As String = ""
            Dim mystrHeader As String = ""
            Dim iHeader As Integer = 0
            Dim j As Integer = 1
            'Console.WriteLine("Create Textfile: " & MyFileName & vbCrLf)
            'Using sw As StreamWriter = New StreamWriter(MyFileName)

            While reader.Read()

                mystr = ""

                For i = 0 To reader.FieldCount - 1
                    'Console.WriteLine("{0} = {1}", reader.GetName(i), reader.GetValue(i))
                    If GetAppKey("HEADER") = "Y" And iHeader = 0 Then
                        If i < reader.FieldCount - 1 Then
                            mystrHeader = reader.GetName(i)
                        Else
                            mystrHeader = reader.GetName(i)
                        End If
                        xlWorkSheet.Cells(1, i + 1) = mystrHeader
                    End If
                    If i < reader.FieldCount - 1 Then
                        mystr = mystr & reader.GetValue(i)
                    Else
                        mystr = mystr & reader.GetValue(i)
                    End If
                    'xlWorkSheet.Columns(1 + i) = mystrHeader 
                    j = j + 1
                    xlWorkSheet.Cells(j, i + 2) = mystr
                Next

                If GetAppKey("HEADER") = "Y" And iHeader = 0 Then
                    iHeader = 1
                    xlWorkSheet.Cells(1, i + 1) = mystrHeader
                    'sw.WriteLine(mystrHeader)
                End If

                'MsgBox(mystr)

                'sw.WriteLine(mystr)


            End While
            xlWorkSheet.SaveAs("c:\SAPM\KB\vbexcel.xlsx")
            xlWorkBook.Close()
            xlApp.Quit()

            releaseObject(xlApp)
            releaseObject(xlWorkBook)
            releaseObject(xlWorkSheet)


            'sw.Close()
            'End Using
            reader.Close()
            conn.Close()
            Console.WriteLine(myText & "Connection closed." & vbCrLf)

        Catch ex As TdException
            Console.WriteLine(myText & "Error: " & ex.ToString & vbCrLf)
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
    Public Sub Export_Only_Cabang(ByRef mystr As String, ByRef sNamaFileZip As String, ByRef sPassword As String)
        'MsgBox(mystr)


        'Set up connection string
        Dim myText As String = "Connecting to database using teradata" & vbCrLf & vbCrLf
        Dim strConn As String = ""
        Dim strConnAttr As String = ""
        Dim strConnVal As String = ""
        Dim connectionString As String = GetAppKey("CONN_STR")
        Dim query As String
        '= (GetAppKey("QUERY2") & mystr)


        Dim args_query(1) As String
        args_query(0) = mystr
        query = QueryBuilder(args_query,
                             (GetAppKey("QUERY2")))

        'MsgBox(query)
        'End

        'Dim abc As String
        'abc = query "where region_kelolaan = '" & mystr & "'"
        '    "Select * from " & txtNamaTabel.Text & " " &
        '            "Where Region_Kelolaan = '" & mystr & "'"

        '"Where Region_Kelolaan = '" & mystr & "'"
        'MsgBox(query)
        'End
        Dim conn As TdConnection

        Try

            conn = New TdConnection(connectionString)
            conn.Open()
            Console.WriteLine(myText & "Connection opened, Process Dataset" & vbCrLf)
            'Console.WriteLine(myQuery)

            Dim tout As Integer = CInt(GetAppKey("TIMEOUT"))
            Dim command = New TdCommand(query, conn)
            command.CommandTimeout = tout
            Dim reader As TdDataReader = command.ExecuteReader
            Dim i As Integer
            Dim strDlm As String = GetAppKey("DELIMITER")
            Dim mystr2 As String = ""
            Dim mystrHeader As String = ""
            Dim iHeader As Integer = 0
            'Console.WriteLine("Create Textfile: " & MyFileName & vbCrLf)
            'Using sw As StreamWriter = New StreamWriter(MyFileName)
            Dim a As Integer
            'Dim b As Integer


            a = 1
            While reader.Read()

                mystr2 = ""
                'reader.FieldCount
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
                        mystr2 = mystr2 & reader.GetValue(i) & strDlm
                    Else
                        mystr2 = mystr2 & reader.GetValue(i)
                    End If

                Next
                If GetAppKey("HEADER") = "Y" And iHeader = 0 Then
                    iHeader = 1

                End If

                If Dir(Trim(txtLokasiFolderKB.Text) & "FileExcel\6" & String.Format("{0:00}", CInt(mystr)) & "\" & String.Format("{0:000}", CInt(mystr2)), vbDirectory) = "" Then
                    MkDir(Trim(txtLokasiFolderKB.Text) & "FileExcel\6" & String.Format("{0:00}", CInt(mystr)) & "\" & String.Format("{0:000}", CInt(mystr2)))
                End If
                If Dir(Trim(txtLokasiFolderKB.Text) & "6" & String.Format("{0:00}", CInt(mystr)) & "\" & String.Format("{0:000}", CInt(mystr2)), vbDirectory) = "" Then
                    MkDir(Trim(txtLokasiFolderKB.Text) & "6" & String.Format("{0:00}", CInt(mystr)) & "\" & String.Format("{0:000}", CInt(mystr2)))
                End If

                Dim sNamaFile As String
                'Dim sNamaFileZip As String
                Dim sNamaFileZipFileInfoData As String
                'Dim sPassword As String

                Dim bulan As String
                bulan = Convert_Date_Str2Int(cmbBulan.Text)


                'If a = 0 Then
                '    a = 1


                'End If
                sNamaFile = Replace(Trim(txtLokasiFolderKB.Text) & "FileExcel\6" & Format(mystr, "00") & "\" & Format(mystr2, "000") & "\Cabang_" & mystr & "_" & mystr2 & "_" & mystr2 & "_" & cmbTahun.Text & (Trim(bulan)) & ".xls", " ", "")
                sPassword = "Pwd" & mystr & mystr2 & cmbTahun.Text & (Trim(bulan))
                'MsgBox(sPassword)
                'sNamaFileZip = Replace(Trim(txtLokasiFolderKB.Text) & "6" & String.Format("{0:00}", CInt(mystr)) & "\" & String.Format("{0:000}", CInt(mystr2)) & "\Cabang_" & mystr & "_" & mystr2 & "_" & cmbTahun.Text & (Trim(bulan)) & ".zip", " ", "")
                sNamaFileZip = Trim(txtLokasiFolderKB.Text) & "6" & String.Format("{0:00}", CInt(mystr)) & "\Wilayah_" & mystr & "_" & cmbTahun.Text & Trim(bulan) & "_" & a & ".zip"
                a = a + 1

                'Dim mystr2 As String
                namaCabangKelolaan(mystr, mystr2)
                'MsgBox(sNamaFileZip)
                'zipfilee(sNamaFileZip)

                'yg lama    'sNamaFileZipFileInfoData = Replace("/home/cmm/DATA/KB/6" & Format(sKodeWilayah, "00") & "/" & Format(sKodeCabang, "000") & "/Cabang_" & sKodeWilayah & "_" & sKodeCabang & "_" & sNamaCabang & "_" & cmbTahun.Text & Trim(gfGetParseData(cmbBulan.Text, 2, "||")) & ".zip", " ", "")
                'sNamaFileZipFileInfoData = Replace("Cabang_" & sKodeWilayah & "_" & sKodeCabang & "_" & sNamaCabang & "_" & cmbTahun.Text & Trim(gfGetParseData(cmbBulan.Text, 2, "||")) & ".zip", " ", "")



                'MsgBox(mystr2)


                'If Dir(Trim(txtLokasiFolderKB.Text) & "FileExcel\6" & String.Format("{0:00}", CInt(mystr)), vbDirectory) = "" Then
                '    MkDir(Trim(txtLokasiFolderKB.Text) & "FileExcel\6" & String.Format("{0:00}", CInt(mystr)))
                'End If
                'If Dir(Trim(txtLokasiFolderKB.Text) & "6" & String.Format("{0:00}", CInt(mystr2)), vbDirectory) = "" Then
                '    MkDir(Trim(txtLokasiFolderKB.Text) & "6" & String.Format("{0:00}", CInt(mystr2)))
                'End If

                'sw.WriteLine(mystr)

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

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        'excel()
        'RunSQLReader()
        'Dim acs As String

        'aku(acs As String)

        'MsgBox(acs)
        'Dim kok As String
        'saya(kok)


        'kok = acs
        'Dim bkn As String

        'Dim zxe As String
        Dim abc As String
        abc = ""
        'zxe = ""

        'Export_Only_Cabang(abc, zxe)
        'saya(abc)
        lll(abc)
        MsgBox(abc)
    End Sub

    Private Sub lll(ByRef zxe As String)
        'Dim zxe As String
        saya(zxe)


    End Sub
    Public Sub saya(ByRef zxe As String)
        'Dim zxe As String
        zxe = "berung"
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

    Private Sub namaCabangKelolaan(ByRef mystr As String, ByRef mystr2 As String)
        Dim sNamaFileZip As String
        'Export_Only_Wilayah_UpdateDita(mystr)
        'Export_Only_Cabang(snamafile, mystr2)
        'Set up connection string
        Dim myText As String = "Connecting to database using teradata" & vbCrLf & vbCrLf
        Dim strConn As String = ""
        Dim strConnAttr As String = ""
        Dim strConnVal As String = ""
        Dim connectionString As String = GetAppKey("CONN_STR")
        Dim conn As TdConnection
        Dim query As String
        '= GetAppKey("QUERY3")

        Dim args_query(1) As String
        args_query(0) = mystr
        args_query(1) = mystr2
        query = QueryBuilder(args_query,
                             (GetAppKey("QUERY3")))
        'MsgBox(query)

        Try

            conn = New TdConnection(connectionString)
            conn.Open()
            Console.WriteLine(myText & "Connection opened, Process Dataset" & vbCrLf)
            'Console.WriteLine(myQuery)

            Dim tout As Integer = CInt(GetAppKey("TIMEOUT"))
            Dim command = New TdCommand(query, conn)
            command.CommandTimeout = tout
            Dim reader As TdDataReader = command.ExecuteReader
            Dim i As Integer
            Dim strDlm As String = GetAppKey("DELIMITER")
            Dim mystr3 As String = ""
            Dim mystrHeader As String = ""
            Dim iHeader As Integer = 0
            'Console.WriteLine("Create Textfile: " & MyFileName & vbCrLf)
            'Using sw As StreamWriter = New StreamWriter(MyFileName)


            While reader.Read()

                mystr3 = ""
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
                        mystr3 = mystr3 & reader.GetValue(i) & strDlm
                    Else
                        mystr3 = mystr3 & reader.GetValue(i)
                    End If
                Next

                If GetAppKey("HEADER") = "Y" And iHeader = 0 Then
                    iHeader = 1
                    'sw.WriteLine(mystrHeader)
                End If
                Dim bulan As String
                bulan = Convert_Date_Str2Int(cmbBulan.Text)

                'If Dir(Trim(txtLokasiFolderKB.Text) & "6" & String.Format("{0:00}", CInt(mystr)) & "\" & String.Format("{0:000}", CInt(mystr2)), vbDirectory) = "" Then
                '    MkDir(Trim(txtLokasiFolderKB.Text) & "6" & String.Format("{0:00}", CInt(mystr)) & "\" & String.Format("{0:000}", CInt(mystr2)))
                'End If
                'Dim sNamaFileZip As String
                sNamaFileZip = Trim(txtLokasiFolderKB.Text) & "6" & String.Format("{0:00}", CInt(mystr)) & "\" & String.Format("{0:000}", CInt(mystr2)) & "\Cabang_" & mystr & "_" & mystr2 & "_" & mystr3 & "_" & cmbTahun.Text & Trim(bulan) & ".zip"
                'MsgBox(sNamaFileZip)
                zipnamacabangfilee(sNamaFileZip)
                'Dim ts As System.IO.StreamWriter
                'ts = My.Computer.FileSystem.OpenTextFileWriter(Trim(txtLokasiFolderKB.Text) & "coba.txt", True)
                'Dim fso As New Object

                'ts.WriteLine(mystr3)
                'ts.Close()

                'MsgBox("masuk")
                'MsgBox(query)
            End While

            'sw.Close()
            'End Using
            reader.Close()
            conn.Close()
            Console.WriteLine(myText & "Connection closed." & vbCrLf)

        Catch ex As TdException
            Console.WriteLine(myText & "Error: " & ex.ToString & vbCrLf)
        End Try

    End Sub

    Private Sub IsiBNISegmentasiDAT(ByRef mystr As String, ByRef mystr2 As String, ByRef sNamaFileZip As String, Optional ByRef sPassword As String = "")
        Dim fso As New Scripting.FileSystemObject
        Dim ts As Scripting.TextStream
        Dim filenya As Scripting.File

        filenya = fso.GetFile(Trim(txtLokasiFolderKB.Text) & "bnisegmentasi.dat")
        ts = filenya.OpenAsTextStream(Scripting.IOMode.ForAppending)

        'ts.WriteLine "6" & Format(sKodeWilayah, "00") & "|" & Format(sKodeCabang, "000") & "|" & sNamaFileZip & " & _ & a &|" & sPassword
        ts.WriteLine("6" & String.Format("{0:00}", CInt(mystr)) & "|" & String.Format("{0:0}", CInt(mystr2)) & "|" & sNamaFileZip & "|" & sPassword)
        'ts.WriteLine "Format(sKodeWilayah, "00") & "|" & sKodeCabang & "|" & sNamaFileZip & "|" & sPassword 'FORMAT LAMA

    End Sub

    Private Sub zip_Click(sender As Object, e As EventArgs) Handles zip.Click
        'cobazip()
        'RunSQLReader()
        textt()


        'Dim mystr As String
        'Dim mystr2 As String
        'namaCabangKelolaan(mystr, mystr2)

        '        Dim file As System.IO.StreamWriter
        '        file = My.Computer.FileSystem.OpenTextFileWriter(txtPathLogFile.Text, True)
        '        file.WriteLine("SerialNo|DateTime|Result|StationName|Model|")
        '        file.Close()

        '        My.Computer.FileSystem.WriteAllText("C:\SAPM\KB",
        '"This is new text to be added.txt", True)

        'Dim fso As New Object
        'Dim ts As System.IO.StreamWriter
        'ts = My.Computer.FileSystem.OpenTextFileWriter(Trim(txtLokasiFolderKB.Text) & "coba.txt", True)
        'ts.WriteLine("")
        'ts.Close()

        'Dim zip As ZipFile




        'Dim ZipToUnpack As String = "C1P3SML.zip"
        'Dim zip As ZipFile
        ''Dim zip As ZipFile = New Zi
        'zip.
        'zip.Encryption = EncryptionAlgorithm.WinZipAes256
        'zip.Password = "123"
        'zip.AddFile("C:\Users\Shahan\Desktop\poetry.B01", "")
        'zip.AddFile("C:\Users\Shahan\Desktop\poetry.inp", "")
        'zip.Save("C:\Users\Shahan\Desktop\Zippedfile.zip")
        'zip.Dispose()


        '        ZipFile.CreateFromDirectory("C:\DataSAPM\",
        '                              "C:\SAPM\KB\tes2.zip",
        '                              CompressionLevel.Optimal,
        ')
        '        ZipFile.Open("C:\SAPM\KB\tes2.zip", open)
        '        Dim sandy As String
        '        sandy = ZipFile.Open("C:\SAPM\KB\tes2.zip", OpenAccess.Read.Read)


        'zipfilee()
        'If File.Exists("C:\DataSAPM\test.txt") Then
        '    ZipFile.CreateFromDirectory("C:\DataSAPM\",
        '                       "c:\SAPM\KB\tes.zip",
        '                       CompressionLevel.Optimal,
        '                       False)
        '    MsgBox("macuk")
        'Else
        '    MsgBox("kosong")
        'End If
        'Dim sNamaFileZip As String
        'zipSandy(sNamaFileZip)

        'MsgBox(sNamaFileZip)

        'If File.Exists("C:\DataSAPM\test.txt") Then
        '    ZipFile.CreateFromDirectory("C:\DataSAPM\",
        '                       "c:\SAPM\KB\tes.zip",
        '                       CompressionLevel.Optimal,
        '                       True)
        'Else
        '    MsgBox("kosong")
        'End If
    End Sub

    Private Sub zipfilee(ByRef sNamaFileZip As String)
        Dim mystr As String
        Dim sPassword As String
        'Dim sNamaFileZip As String
        Export_Only_Cabang(mystr, sNamaFileZip, sPassword)

        MsgBox(sNamaFileZip)
        ZipFile.CreateFromDirectory("C:\DataSAPM\",
                               sNamaFileZip,
                               CompressionLevel.Optimal,
                               False)
        Dim sNamaFileAsli As String = "C:\DataSAPM\test.txt"
        'Dim sPassword As String = "123"
        Dim sNamaFileZipa As String = "c:\SAPM\KB\tes.zip"
        Dim tzt As String = txtLokasiFileWinrar.Text

        gpCompressFileToZip(tzt, sNamaFileAsli, sPassword, sNamaFileZipa)
        'If File.Exists("C:\DataSAPM\test.txt") Then
        '    ZipFile.CreateFromDirectory("C:\DataSAPM\",
        '                       "c:\SAPM\KB\tes.zip",
        '                       CompressionLevel.Optimal,
        '                       False)
        '    MsgBox("macuk")
        'Else
        '    MsgBox("kosong")
        'End If
    End Sub

    Private Sub cobazip()
        Dim sNamaFileAsli As String = "C:\DataSAPM\test.txt"
        Dim sPassword As String = "123"
        Dim sNamaFileZipa As String = "c:\SAPM\KB\tes.zip"
        Dim tzt As String = txtLokasiFileWinrar.Text

        gpCompressFileToZip(tzt, sNamaFileAsli, sPassword, sNamaFileZipa)
    End Sub
    Public Sub gpCompressFileToZip(ByRef sLokasiWinrar As String, ByRef sNamaFileAsli As String, ByRef sPassword As String, ByRef sNamaFileZip As String)
        Dim mystr As String
        'Dim sPassword As String
        'Dim sNamaFileZip As String
        Export_Only_Cabang(mystr, sNamaFileZip, sPassword)
        Shell(sLokasiWinrar & " a -p" & sPassword & " " & sNamaFileZip & " " & sNamaFileAsli)
    End Sub


    Private Sub zipnamacabangfilee(ByRef sNamaFileZip As String)
        'Using ZipArchive = ZipFile.Open("c:\SAPM\KB\tes.zip", ZipArchiveMode.Update)
        '    ZipArchive.
        'End Using
        'Dim mystr As String
        'Dim sNamaFileZip As String
        'Export_Only_Cabang(mystr, sNamaFileZip)

        'MsgBox(sNamaFileZip)
        Dim sNamaFileAsli As String = "C:\DataSAPM\test.txt"
        Dim sPassword As String = "123"
        Dim sNamaFileZipa As String = "c:\SAPM\KB\tes.zip"
        Dim tzt As String = txtLokasiFileWinrar.Text

        gpCompressFileToZip(tzt, sNamaFileAsli, sPassword, sNamaFileZipa)

        ZipFile.CreateFromDirectory("C:\DataSAPM\",
                               sNamaFileZip,
                               CompressionLevel.Optimal,
                               False)
    End Sub

    Private Sub textt()
        'Set up connection string
        Dim myText As String = "Connecting to database using teradata" & vbCrLf & vbCrLf
        Dim strConn As String = ""
        Dim strConnAttr As String = ""
        Dim strConnVal As String = ""
        Dim connectionString As String = GetAppKey("CONN_STR")
        Dim QUERY As String = GetAppKey("QUERYTXTCABANG")
        Dim conn As TdConnection

        Try

            conn = New TdConnection(connectionString)
            conn.Open()
            Console.WriteLine(myText & "Connection opened, Process Dataset" & vbCrLf)
            'Console.WriteLine(myQuery)

            Dim tout As Integer = CInt(GetAppKey("TIMEOUT"))
            Dim command = New TdCommand(QUERY, conn)
            command.CommandTimeout = tout
            Dim reader As TdDataReader = command.ExecuteReader
            Dim i As Integer
            Dim strDlm As String = GetAppKey("DELIMITER")
            Dim mystr As String = ""
            Dim mystrHeader As String = ""
            Dim iHeader As Integer = 0
            'Console.WriteLine("Create Textfile: " & MyFileName & vbCrLf)
            'Using sw As StreamWriter = New StreamWriter(MyFileName)

            While reader.Read()

                mystr = ""
                'reader.FieldCount
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
                Dim ts As System.IO.StreamWriter
                ts = My.Computer.FileSystem.OpenTextFileWriter(Trim(txtLokasiFolderKB.Text) & "bnisegmentasi.txt", True)

                If GetAppKey("HEADER") = "Y" And iHeader = 0 Then
                    iHeader = 1
                    ts.WriteLine(mystrHeader)
                End If

                ts.WriteLine(mystr)

                'ts.WriteLine(mystrHeader)
                ts.Close()

            End While

            'sw.Close()
            'End Using
            reader.Close()
            conn.Close()
            Console.WriteLine(myText & "Connection closed." & vbCrLf)

        Catch ex As TdException
            Console.WriteLine(myText & "Error: " & ex.ToString & vbCrLf)
        End Try

    End Sub


    'Function zipSandy(sNamaFileZip As String)
    '    If File.Exists("C:\DataSAPM\20191231_with_dormanbnisegmentasi.dat") Then
    '        ZipFile.CreateFromDirectory("C:\SAPM\KB\FileExcel",
    '                           "c:\SAPM\KB\tes.zip",
    '                           CompressionLevel.Optimal,
    '                           True)
    '    Else
    '        MsgBox("kosong")
    '    End If

    'End Function

End Class
