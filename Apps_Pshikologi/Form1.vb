Imports System.IO
Imports System.Data.OleDb
Imports CrystalDecisions.CrystalReports.Engine

Imports Microsoft.Office.Interop
Imports System.Runtime.InteropServices.Marshal
Public Class Form1
    Dim nama As String = Nothing
    Dim extensi As String = Nothing
    Dim myLokasi As String = Application.StartupPath

    Dim konekExcel As OleDbConnection
    Dim komen As OleDbCommand = New OleDbCommand()
    Dim baca As OleDbDataAdapter = New OleDbDataAdapter()

    Dim database = "db_ps.mdb"

    Dim konekAccess As OleDbConnection
    Dim komenAccess As New OleDbCommand
    Dim bacaAccess As OleDbDataReader
    Dim data As DataSet = New DataSet()

    Dim dataShetPsikoGram As New dtPsikogram
    Dim dataTablePsikogram As DataTable = dataShetPsikoGram.Tables("DataTable1")

    Dim id_ist As String

    Sub clear_database()
        Try
            konekAccess.Open()
            komenAccess.Connection = konekAccess
            komenAccess.CommandText = "DELETE FROM nilai"
            komenAccess.ExecuteNonQuery()
            konekAccess.Close()
        Catch ex As Exception
            konekAccess.Close()
            MessageBox.Show(ex.Message, "Delete")
        End Try
    End Sub

    Sub load_judul()
        Try
            konekAccess.Open()
            komenAccess.Connection = konekAccess
            komenAccess.CommandText = "SELECT * FROM nilai"
            komenAccess.ExecuteNonQuery()

            bacaAccess = komenAccess.ExecuteReader
            bacaAccess.Read()

            TextBox1.Text = bacaAccess(18)
            TextBox2.Text = bacaAccess(34)
            TextBox3.Text = bacaAccess(35)
            konekAccess.Close()
        Catch ex As Exception
            konekAccess.Close()
            MsgBox(ex.Message)
        End Try
    End Sub

    Sub load_data_access()
        Try
            konekAccess.Open()
            komenAccess.Connection = konekAccess
            komenAccess.CommandText = "SELECT * FROM nilai"
            komenAccess.ExecuteNonQuery()

            ListView1.Items.Clear()
            dataTablePsikogram.Rows.Clear()

            bacaAccess = komenAccess.ExecuteReader
            While bacaAccess.Read
                dataTablePsikogram.Rows.Add(bacaAccess(0).ToString, bacaAccess(1).ToString, _
                                            bacaAccess(2).ToString, bacaAccess(3).ToString, _
                                            bacaAccess(4).ToString, bacaAccess(5).ToString, _
                                            bacaAccess(6).ToString, bacaAccess(7).ToString, _
                                            bacaAccess(8).ToString, bacaAccess(9).ToString, _
                                            bacaAccess(10).ToString, bacaAccess(11).ToString, _
                                            bacaAccess(12).ToString, bacaAccess(13).ToString, _
                                            bacaAccess(14).ToString, bacaAccess(15).ToString, _
                                            bacaAccess(16).ToString, bacaAccess(17).ToString, bacaAccess(18).ToString, _
                                            bacaAccess(19).ToString, bacaAccess(20).ToString, bacaAccess(21).ToString, _
                                            bacaAccess(22).ToString, bacaAccess(23).ToString, bacaAccess(24).ToString, _
                                            bacaAccess(25).ToString, bacaAccess(26).ToString, bacaAccess(27).ToString, _
                                            bacaAccess(28).ToString, bacaAccess(29).ToString, bacaAccess(30).ToString, _
                                            bacaAccess(31).ToString, bacaAccess(32).ToString, bacaAccess(33).ToString, _
                                            bacaAccess(34).ToString, bacaAccess(35).ToString, bacaAccess(36).ToString, bacaAccess(37).ToString, bacaAccess(38).ToString)

                With ListView1.Items.Add(bacaAccess(0).ToString)
                    .SubItems.Add(bacaAccess(1).ToString)
                    .SubItems.Add(bacaAccess(2).ToString)
                    .SubItems.Add(bacaAccess(3).ToString)
                    .SubItems.Add(bacaAccess(4).ToString)
                    .SubItems.Add(bacaAccess(5).ToString)
                    .SubItems.Add(bacaAccess(6).ToString)
                    .SubItems.Add(bacaAccess(7).ToString)
                    .SubItems.Add(bacaAccess(8).ToString)
                    .SubItems.Add(bacaAccess(9).ToString)
                    .SubItems.Add(bacaAccess(10).ToString)
                    .SubItems.Add(bacaAccess(11).ToString)
                    .SubItems.Add(bacaAccess(12).ToString)
                    .SubItems.Add(bacaAccess(13).ToString)
                    .SubItems.Add(bacaAccess(14).ToString)
                    .SubItems.Add(bacaAccess(15).ToString)
                    .SubItems.Add(bacaAccess(16).ToString)
                    .SubItems.Add(bacaAccess(17).ToString)
                    .SubItems.Add(bacaAccess(18).ToString)
                    .SubItems.Add(bacaAccess(19).ToString)
                    .SubItems.Add(bacaAccess(20).ToString)
                    .SubItems.Add(bacaAccess(21).ToString)
                    .SubItems.Add(bacaAccess(22).ToString)
                    .SubItems.Add(bacaAccess(23).ToString)
                    .SubItems.Add(bacaAccess(24).ToString)
                    .SubItems.Add(bacaAccess(25).ToString)
                    .SubItems.Add(bacaAccess(26).ToString)
                    .SubItems.Add(bacaAccess(27).ToString)
                    .SubItems.Add(bacaAccess(28).ToString)
                    .SubItems.Add(bacaAccess(29).ToString)
                    .SubItems.Add(bacaAccess(30).ToString)
                    .SubItems.Add(bacaAccess(31).ToString)
                    .SubItems.Add(bacaAccess(32).ToString)
                    .SubItems.Add(bacaAccess(33).ToString)
                    .SubItems.Add(bacaAccess(34).ToString)
                    .SubItems.Add(bacaAccess(35).ToString)
                    .SubItems.Add(bacaAccess(36).ToString)
                    .SubItems.Add(bacaAccess(37).ToString)
                    .SubItems.Add(bacaAccess(38).ToString)
                End With
            End While
            konekAccess.Close()
        Catch ex As Exception
            konekAccess.Close()
            MessageBox.Show(ex.Message, "Delete")
        End Try
    End Sub

    Sub load_data_excel()
        konekExcel = New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + Application.StartupPath + "\" + nama + ";Extended Properties=Excel 12.0;")

        Try
            konekExcel.Open()
            komen.Connection = konekExcel
            komen.CommandText = "select * from [TEMPORARYOUT$]"

            baca.SelectCommand = komen
            baca.Fill(data)
            ListView1.Items.Clear()
            dataTablePsikogram.Rows.Clear()

            'konekAccess = New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + Application.StartupPath + "\db_ps.accdb;")
            For x As Integer = 0 To data.Tables(0).Rows.Count - 1
                Try
                    konekAccess.Open()
                    komenAccess.Connection = konekAccess
                    Dim query As String = "INSERT INTO nilai VALUES(" + data.Tables(0).Rows(x).ItemArray(0).ToString + ",'" + data.Tables(0).Rows(x).ItemArray(1).ToString + "'," + _
                                                "'" + data.Tables(0).Rows(x).ItemArray(2).ToString + "','" + data.Tables(0).Rows(x).ItemArray(3).ToString + "'," + _
                                                "'" + data.Tables(0).Rows(x).ItemArray(4).ToString + "','" + data.Tables(0).Rows(x).ItemArray(5).ToString + "'," + _
                                                "'" + data.Tables(0).Rows(x).ItemArray(6).ToString + "','" + data.Tables(0).Rows(x).ItemArray(7).ToString + "'," + _
                                                "'" + data.Tables(0).Rows(x).ItemArray(29).ToString + "','" + data.Tables(0).Rows(x).ItemArray(30).ToString + "'," + _
                                                "'" + data.Tables(0).Rows(x).ItemArray(31).ToString + "','" + data.Tables(0).Rows(x).ItemArray(32).ToString + "'," + _
                                                "'" + data.Tables(0).Rows(x).ItemArray(33).ToString + "','" + data.Tables(0).Rows(x).ItemArray(34).ToString + "'," + _
                                                "'" + data.Tables(0).Rows(x).ItemArray(35).ToString + "','" + data.Tables(0).Rows(x).ItemArray(36).ToString + "'," + _
                                                "'" + data.Tables(0).Rows(x).ItemArray(37).ToString + "','" + data.Tables(0).Rows(x).ItemArray(39).ToString + "'," + _
                                                "'" + TextBox1.Text + "/" + data.Tables(0).Rows(x).ItemArray(1).ToString + "','','','','','','','','','','','','','','','','','','','','')"
                    komenAccess.CommandText = query
                    komenAccess.ExecuteNonQuery()
                    konekAccess.Close()
                    Application.DoEvents()
                Catch ex As Exception
                    konekAccess.Close()
                    MessageBox.Show(ex.Message)
                End Try
            Next

            konekExcel.Close()
        Catch ex As Exception
            konekExcel.Close()
            MessageBox.Show(ex.Message, "Kesalahan")
        End Try
    End Sub

    Sub load_data_excel_2()
        konekExcel = New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + Application.StartupPath + "\" + nama + ";Extended Properties=Excel 12.0;")

        Try
            konekExcel.Open()
            komen.Connection = konekExcel
            komen.CommandText = "select * from [MBTI$] where NAMA <> ''"

            baca.SelectCommand = komen
            baca.Fill(data)
            ListView1.Items.Clear()
            dataTablePsikogram.Rows.Clear()


            'konekAccess = New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + Application.StartupPath + "\db_ps.accdb;")
            For x As Integer = 0 To data.Tables(0).Rows.Count - 1
                Try

                    Dim ekstrovet As Integer = data.Tables(0).Rows(x).ItemArray(1)
                    Dim instrovet As Integer = data.Tables(0).Rows(x).ItemArray(2)
                    Dim praktis As Integer = data.Tables(0).Rows(x).ItemArray(3)
                    Dim inovatif As Integer = data.Tables(0).Rows(x).ItemArray(4)
                    Dim rasional As Integer = data.Tables(0).Rows(x).ItemArray(5)
                    Dim bijaksana As Integer = data.Tables(0).Rows(x).ItemArray(6)
                    Dim terencana As Integer = data.Tables(0).Rows(x).ItemArray(7)
                    Dim spontan As Integer = data.Tables(0).Rows(x).ItemArray(8)
                    Dim status1 As String
                    Dim status2 As String
                    Dim status3 As String
                    Dim status4 As String
                    'MsgBox("" & ekstrovet & " " & instrovet & " " & praktis & " " & inovatif & " " & rasional & " " & bijaksana & " " & terencana & " " & spontan)
                    ekstrovet = (ekstrovet / (ekstrovet + instrovet)) * 100
                    instrovet = 100 - ekstrovet
                    praktis = (praktis / (praktis + inovatif)) * 100
                    inovatif = 100 - praktis
                    rasional = (rasional / (rasional + bijaksana)) * 100
                    bijaksana = 100 - rasional
                    terencana = (terencana / (terencana + spontan)) * 100
                    spontan = 100 - terencana

                    If ekstrovet < instrovet Then
                        status1 = "INTROVERT"
                    ElseIf ekstrovet = instrovet Then
                        status1 = "SEIMBANG"
                    Else
                        status1 = "EKSTROVERT"
                    End If

                    If praktis < inovatif Then
                        status2 = "INOVATIF"
                    ElseIf praktis = inovatif Then
                        status2 = "SEIMBANG"
                    Else
                        status2 = "PRAKTIS"
                    End If

                    If rasional < bijaksana Then
                        status3 = "BIJAKSANA"
                    ElseIf rasional = bijaksana Then
                        status3 = "SEIMBANG"
                    Else
                        status3 = "RASIONAL"
                    End If

                    If terencana < spontan Then
                        status4 = "TERENCANA"
                    ElseIf terencana = spontan Then
                        status4 = "SEIMBANG"
                    Else
                        status4 = "SPONTAN"
                    End If

                    konekAccess.Open()
                    komenAccess.Connection = konekAccess
                    Dim query As String = "UPDATE nilai SET EXTROVERT='" + Convert.ToString(ekstrovet) + "', INTROVERT='" +
                        Convert.ToString(instrovet) + "', PRAKTIS='" + Convert.ToString(praktis) + "', INOVATIF='" +
                        Convert.ToString(inovatif) + "', RASIONAL='" + Convert.ToString(rasional) + "', BIJAKSANA='" +
                        Convert.ToString(bijaksana) + "', TERENCANA='" + Convert.ToString(terencana) + "', SPONTAN = '" +
                        Convert.ToString(spontan) + "', STATUS1='" + status1 + "', STATUS2='" + status2 + "', STATUS3 ='" +
                        status3 + "', STATUS4='" + status4 + "' WHERE NAMA = '" + data.Tables(0).Rows(x).ItemArray(0).ToString + "'; "
                    'MessageBox.Show(query)
                    komenAccess.CommandText = query
                    komenAccess.ExecuteNonQuery()
                    konekAccess.Close()
                    Application.DoEvents()
                Catch ex As Exception
                    konekAccess.Close()
                    MessageBox.Show(ex.Message, "Insert MBTI to Access")
                End Try
            Next

            konekExcel.Close()
        Catch ex As Exception
            konekExcel.Close()
            MessageBox.Show(ex.Message, "Kesalahan MBTI")
        End Try

        'konekExcel = New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + Application.StartupPath + "\" + nama + ";Extended Properties=Excel 12.0;")

        komen = New OleDbCommand()
        Try
            konekExcel.Open()
            komen.Connection = konekExcel
            komen.CommandText = "select * from [NLP$]"

            data.Reset()
            baca.SelectCommand = komen
            baca.Fill(data)
            ListView1.Items.Clear()
            dataTablePsikogram.Rows.Clear()

            'MsgBox(data.Tables(0).Rows.Count)
            For x As Integer = 0 To data.Tables(0).Rows.Count - 1
                Try
                    'MsgBox(data.Tables(0).Rows(x).ItemArray(2).ToString)
                    Dim visual As Integer = data.Tables(0).Rows(x).ItemArray(1)
                    Dim auditori As Integer = data.Tables(0).Rows(x).ItemArray(2)
                    Dim kinestetik As Integer = data.Tables(0).Rows(x).ItemArray(3)

                    'MsgBox(visual & " " & auditori & " " & kinestetik)
                    visual = visual * 10
                    auditori = auditori * 10
                    kinestetik = kinestetik * 10
                    konekAccess.Open()
                    komenAccess.Connection = konekAccess
                    Dim query As String = "UPDATE nilai SET VISUAL = '" + visual.ToString() + "', AUDITORI = '" +
                        auditori.ToString() + "', KINESTETIK='" + kinestetik.ToString() + "' WHERE NAMA = '" + data.Tables(0).Rows(x).ItemArray(0).ToString() + "'; "
                    'MessageBox.Show(query)
                    komenAccess.CommandText = query
                    komenAccess.ExecuteNonQuery()
                    konekAccess.Close()
                    Application.DoEvents()
                Catch ex As Exception
                    konekAccess.Close()
                    MessageBox.Show(ex.Message, "Insert To Access NLP")
                End Try
            Next

            konekExcel.Close()
        Catch ex As Exception
            konekExcel.Close()
            MessageBox.Show(ex.Message, "Kesalahan NLP")
        End Try

        konekExcel = New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + Application.StartupPath + "\" + nama + ";Extended Properties=Excel 12.0;")
        Try
            konekExcel.Open()
            komen.Connection = konekExcel
            komen.CommandText = "select * from [RMIB$]"

            data.Reset()
            baca.SelectCommand = komen
            baca.Fill(data)
            ListView1.Items.Clear()
            dataTablePsikogram.Rows.Clear()

            For x As Integer = 0 To data.Tables(0).Rows.Count - 1
                Try
                    'MessageBox.Show(Convert.ToString(x))
                    konekAccess.Open()
                    komenAccess.Connection = konekAccess
                    Dim query As String = "UPDATE nilai SET SETA='" + data.Tables(0).Rows(x).ItemArray(1).ToString() + "', SETB = '" +
                        data.Tables(0).Rows(x).ItemArray(2).ToString() + "', SETC = '" + data.Tables(0).Rows(x).ItemArray(3).ToString() + "' WHERE NAMA = '" +
                        data.Tables(0).Rows(x).ItemArray(0).ToString() + "'; "
                    'MessageBox.Show(query)
                    komenAccess.CommandText = query
                    komenAccess.ExecuteNonQuery()
                    konekAccess.Close()
                    Application.DoEvents()
                Catch ex As Exception
                    konekAccess.Close()
                    MessageBox.Show(ex.Message)
                End Try
            Next

            konekExcel.Close()
        Catch ex As Exception
            konekExcel.Close()
            MessageBox.Show(ex.Message, "Kesalahan")
        End Try

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If TextBox1.Text <> "" Then
            Dim pathFile As String = Nothing
            OpenFileDialog1.Filter = "Excel 97-2003(*.xls)|*.xls|Excel Workbook(*.xlsx)|*.xlsx"
            OpenFileDialog1.FileName = ""
            If OpenFileDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
                pathFile = OpenFileDialog1.FileName
                nama = pathFile.Substring(pathFile.LastIndexOf("\") + 1)
                extensi = pathFile.Substring(pathFile.LastIndexOf(".") + 1)

                My.Settings.NamaFile = Split(nama, ".")(0)
                My.Settings.Save()

                If (File.Exists(myLokasi + "\" + nama)) Then
                    File.Delete(myLokasi + "\" + nama)
                    File.Copy(pathFile, myLokasi + "\" + nama)
                    clear_database()
                    load_data_excel()
                    load_data_access()
                Else
                    File.Copy(pathFile, myLokasi + "\" + nama)
                    clear_database()
                    load_data_excel()
                    load_data_access()
                End If
            End If
        Else
            MessageBox.Show("Mohon isi kode registrasi")
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Cursor = Cursors.AppStarting
        Dim reportDoc As ReportDocument
        reportDoc = New rPsikogram
        reportDoc.SetDataSource(dataTablePsikogram)
        Form2.CrystalReportViewer1.Refresh()
        Form2.CrystalReportViewer1.ReportSource = reportDoc
        Form2.Show()
        Me.Cursor = Cursors.Default
    End Sub

    Private Sub nudEkstrovert_ValueChanged(sender As Object, e As EventArgs) Handles nudEkstrovert.ValueChanged, nudIntrovert.ValueChanged, _
                                                                                        nudPraktis.ValueChanged, nudInovatif.ValueChanged, _
                                                                                        nudRasional.ValueChanged, nudBijaksana.ValueChanged, _
                                                                                        nudTerencana.ValueChanged, nudSpontan.ValueChanged, _
                                                                                        nudAuditori.ValueChanged, nudVisual.ValueChanged, nudKinestik.ValueChanged
        Dim obj As NumericUpDown = sender
        'CMP
        If nudEkstrovert.Name = obj.Name Then
            nudIntrovert.Value = 100 - nudEkstrovert.Value
            setLabel(lbSTATUS1, nudEkstrovert.Value, "EKSTROVERT", "INTROVERT")
        End If
        If nudIntrovert.Name = obj.Name Then
            nudEkstrovert.Value = 100 - nudIntrovert.Value
            setLabel(lbSTATUS1, nudIntrovert.Value, "INTROVERT", "EKSTROVERT")
        End If

        'CMI
        If nudPraktis.Name = obj.Name Then
            nudInovatif.Value = 100 - nudPraktis.Value
            setLabel(lbSTATUS2, nudPraktis.Value, "PRAKTIS", "INOVATIF")
        End If
        If nudInovatif.Name = obj.Name Then
            nudPraktis.Value = 100 - nudInovatif.Value
            setLabel(lbSTATUS2, nudInovatif.Value, "INOVATIF", "PRAKTIS")
        End If

        'CMK
        If nudRasional.Name = obj.Name Then
            nudBijaksana.Value = 100 - nudRasional.Value
            setLabel(lbSTATUS3, nudRasional.Value, "RASIONAL", "BIJAKSANA")
        End If
        If nudBijaksana.Name = obj.Name Then
            nudRasional.Value = 100 - nudBijaksana.Value
            setLabel(lbSTATUS3, nudBijaksana.Value, "BIJAKSANA", "RASIONAL")
        End If

        'CPH
        If nudTerencana.Name = obj.Name Then
            nudSpontan.Value = 100 - nudTerencana.Value
            setLabel(lbSTATUS4, nudTerencana.Value, "TERENCANA", "SPONTAN")
        End If
        If nudSpontan.Name = obj.Name Then
            nudTerencana.Value = 100 - nudSpontan.Value
            setLabel(lbSTATUS4, nudSpontan.Value, "SPONTAN", "TERENCANA")
        End If

        If nudVisual.Name = obj.Name Or nudAuditori.Name = obj.Name Or nudKinestik.Name = obj.Name Then
            lbTotalVAK.Text = nudVisual.Value + nudAuditori.Value + nudKinestik.Value & "%"
            If Val(lbTotalVAK.Text) > 100 Then
                lbTotalVAK.ForeColor = Color.Red
            Else
                lbTotalVAK.ForeColor = Color.Black
            End If
        End If
    End Sub

    Dim sisa As Integer = 0
    Sub setLabel(lb As Label, value As Integer, max As String, min As String)
        If value = 50 Then
            lb.Text = "SEIMBANG"
        ElseIf value > 50 Then
            lb.Text = max
        Else
            lb.Text = min
        End If
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        konekAccess = New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + Application.StartupPath + "\" + database + ";")

        load_data_access()
        load_judul()
    End Sub

    Private Sub ListView1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListView1.DoubleClick
        Try
            With ListView1.SelectedItems(0)
                id_ist = .SubItems(0).Text
                nudEkstrovert.Value = Val(.SubItems(19).Text)
                nudIntrovert.Value = Val(.SubItems(20).Text)
                lbSTATUS1.Text = .SubItems(21).Text

                nudPraktis.Value = Val(.SubItems(22).Text)
                nudInovatif.Value = Val(.SubItems(23).Text)
                lbSTATUS2.Text = .SubItems(24).Text

                nudRasional.Value = Val(.SubItems(25).Text)
                nudBijaksana.Value = Val(.SubItems(26).Text)
                lbSTATUS3.Text = .SubItems(27).Text

                nudTerencana.Value = Val(.SubItems(28).Text)
                nudSpontan.Value = Val(.SubItems(29).Text)
                lbSTATUS4.Text = .SubItems(30).Text

                nudVisual.Value = Val(.SubItems(31).Text)
                nudAuditori.Value = Val(.SubItems(32).Text)
                nudKinestik.Value = Val(.SubItems(33).Text)

                NumericUpDown1.Value = Val(.SubItems(36).Text)
                NumericUpDown2.Value = Val(.SubItems(37).Text)
                NumericUpDown3.Value = Val(.SubItems(38).Text)
            End With
        Catch ex As Exception

        End Try
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Try
            konekAccess.Open()
            komenAccess.Connection = konekAccess
            komenAccess.CommandText = "UPDATE nilai SET EXTROVERT='" + String.Format("{0}", nudEkstrovert.Value) + "', INTROVERT='" + _
                String.Format("{0}", nudIntrovert.Value) + "', STATUS1='" + lbSTATUS1.Text + "', PRAKTIS='" + String.Format("{0}", nudPraktis.Value) + _
                "', INOVATIF='" + String.Format("{0}", nudInovatif.Value) + "', STATUS2='" + lbSTATUS2.Text + "', RASIONAL='" + String.Format("{0}", nudRasional.Value) + "', BIJAKSANA='" + _
                String.Format("{0}", nudBijaksana.Value) + "', STATUS3='" + lbSTATUS3.Text + "', TERENCANA='" + String.Format("{0}", nudTerencana.Value) + "', SPONTAN='" + String.Format("{0}", nudSpontan.Value) + _
                "', STATUS4='" + lbSTATUS4.Text + "', VISUAL='" + String.Format("{0}", nudVisual.Value) + "', AUDITORI='" + String.Format("{0}", nudAuditori.Value) + "', KINESTETIK='" + _
                String.Format("{0}", nudKinestik.Value) + "', SETA='" + String.Format("{0}", NumericUpDown1.Value) + "', SETB='" + String.Format("{0}", NumericUpDown2.Value) + "', SETC='" + String.Format("{0}", NumericUpDown3.Value) + "' WHERE ID_IST=" + id_ist
            komenAccess.ExecuteNonQuery()
            konekAccess.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "UPDATE")
        End Try
        load_data_access()
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Me.Cursor = Cursors.AppStarting
        If TextBox2.Text = "" Or TextBox3.Text = "" Then
            MessageBox.Show("Mohon isi judul dan instansi rekap data")
        Else
            Try
                konekAccess.Open()
                komenAccess.Connection = konekAccess
                komenAccess.CommandText = "UPDATE nilai SET INSTANSI='" + TextBox2.Text + "', JUDUL='" + TextBox3.Text + "' WHERE 1"
                komenAccess.ExecuteNonQuery()
                konekAccess.Close()
            Catch ex As Exception
                MessageBox.Show(ex.Message, "Update Judul")
            End Try

            load_data_access()

            Dim reportDoc As ReportDocument
            reportDoc = New rRekap
            reportDoc.SetDataSource(dataTablePsikogram)
            Form2.CrystalReportViewer1.Refresh()
            Form2.CrystalReportViewer1.ReportSource = reportDoc
            Form2.Show()
        End If
        Me.Cursor = Cursors.Default
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Dim pathFile As String = Nothing
        OpenFileDialog1.Filter = "Access 2000-2003(*.mdb)|*.mdb"
        OpenFileDialog1.FileName = ""
        If OpenFileDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
            pathFile = OpenFileDialog1.FileName
            database = pathFile.Substring(pathFile.LastIndexOf("\") + 1)

            konekAccess = New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + pathFile + ";")
            'extensi = pathFile.Substring(pathFile.LastIndexOf(".") + 1)
            load_data_access()
            load_judul()
            'MessageBox.Show(pathFile)
        End If
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Me.Cursor = Cursors.AppStarting
        Dim x As Integer
        Try
            konekAccess.Open()
            komenAccess.Connection = konekAccess
            komenAccess.CommandText = "SELECT COUNT(ID_IST) as jumlah FROM nilai"
            komenAccess.ExecuteNonQuery()
            bacaAccess = komenAccess.ExecuteReader
            bacaAccess.Read()

            x = Val(bacaAccess("jumlah"))
            konekAccess.Close()
        Catch ex As Exception
            konekAccess.Close()
            MessageBox.Show(ex.Message)
        End Try

        For jum As Integer = 1 To x
            Try
                Dim tmp As String = TextBox1.Text + "/" + String.Format("{0:00}", jum)
                konekAccess.Open()
                komenAccess.Connection = konekAccess
                komenAccess.CommandText = "UPDATE nilai SET NOREG='" + tmp + "' WHERE NOMOR='" + String.Format("{0:00}", jum) + "'"
                komenAccess.ExecuteNonQuery()

                konekAccess.Close()
            Catch ex As Exception
                konekAccess.Close()
                MessageBox.Show(ex.Message)
            End Try
        Next
        load_data_access()
        Me.Cursor = Cursors.Default
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        Try
            konekAccess.Open()
            komenAccess.Connection = konekAccess
            komenAccess.CommandText = "UPDATE nilai SET EXTROVERT='" + String.Format("{0}", nudEkstrovert.Value) + "', INTROVERT='" + _
                String.Format("{0}", nudIntrovert.Value) + "', STATUS1='" + lbSTATUS1.Text + "', PRAKTIS='" + String.Format("{0}", nudPraktis.Value) + _
                "', INOVATIF='" + String.Format("{0}", nudInovatif.Value) + "', STATUS2='" + lbSTATUS2.Text + "', RASIONAL='" + String.Format("{0}", nudRasional.Value) + "', BIJAKSANA='" + _
                String.Format("{0}", nudBijaksana.Value) + "', STATUS3='" + lbSTATUS3.Text + "', TERENCANA='" + String.Format("{0}", nudTerencana.Value) + "', SPONTAN='" + String.Format("{0}", nudSpontan.Value) + _
                "', STATUS4='" + lbSTATUS4.Text + "', VISUAL='" + String.Format("{0}", nudVisual.Value) + "', AUDITORI='" + String.Format("{0}", nudAuditori.Value) + "', KINESTETIK='" + _
                String.Format("{0}", nudKinestik.Value) + "', SETA='" + String.Format("{0}", NumericUpDown1.Value) + "', SETB='" + String.Format("{0}", NumericUpDown2.Value) + "', SETC='" + String.Format("{0}", NumericUpDown3.Value) + "' WHERE ID_IST=" + id_ist
            komenAccess.ExecuteNonQuery()
            konekAccess.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "UPDATE")
        End Try
        load_data_access()

        If File.Exists(myLokasi + "\" + My.Settings.NamaFile + ".mdb") Then
            File.Delete(myLokasi + "\" + My.Settings.NamaFile + ".mdb")
            File.Copy(myLokasi + "\" + database, myLokasi + "\" + My.Settings.NamaFile + ".mdb")
            MsgBox("Data Telah disimpan!")
        Else
            File.Copy(myLokasi + "\" + database, myLokasi + "\" + My.Settings.NamaFile + ".mdb")
            MsgBox("Data Telah disimpan!")
        End If
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs)
        If TextBox1.Text <> "" Then
            Dim pathFile As String = Nothing
            OpenFileDialog1.Filter = "Excel 97-2003(*.xls)|*.xls|Excel Workbook(*.xlsx)|*.xlsx"
            OpenFileDialog1.FileName = ""
            If OpenFileDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
                pathFile = OpenFileDialog1.FileName
                nama = pathFile.Substring(pathFile.LastIndexOf("\") + 1)
                extensi = pathFile.Substring(pathFile.LastIndexOf(".") + 1)

                My.Settings.NamaFile = Split(nama, ".")(0)
                My.Settings.Save()

                If (File.Exists(myLokasi + "\" + nama)) Then
                    File.Delete(myLokasi + "\" + nama)
                    File.Copy(pathFile, myLokasi + "\" + nama)
                    'clear_database()
                    load_data_excel_2()
                    load_data_access()
                Else
                    File.Copy(pathFile, myLokasi + "\" + nama)
                    'clear_database()
                    load_data_excel_2()
                    load_data_access()
                End If
            End If
        Else
            MessageBox.Show("Mohon isi kode registrasi")
        End If
    End Sub

    Private Sub Button8_Click_1(sender As Object, e As EventArgs) Handles Button8.Click
        If TextBox1.Text <> "" Then
            Dim pathFile As String = Nothing
            OpenFileDialog1.Filter = "Excel 97-2003(*.xls)|*.xls|Excel Workbook(*.xlsx)|*.xlsx"
            OpenFileDialog1.FileName = ""
            If OpenFileDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
                pathFile = OpenFileDialog1.FileName
                nama = pathFile.Substring(pathFile.LastIndexOf("\") + 1)
                extensi = pathFile.Substring(pathFile.LastIndexOf(".") + 1)

                My.Settings.NamaFile = Split(nama, ".")(0)
                My.Settings.Save()

                If (File.Exists(myLokasi + "\" + nama)) Then
                    File.Delete(myLokasi + "\" + nama)
                    File.Copy(pathFile, myLokasi + "\" + nama)
                    'clear_database()
                    load_data_excel_2()
                    load_data_access()
                Else
                    File.Copy(pathFile, myLokasi + "\" + nama)
                    'clear_database()
                    load_data_excel_2()
                    load_data_access()
                End If
            End If
        Else
            MessageBox.Show("Mohon isi kode registrasi")
        End If
    End Sub
End Class
