Imports System.Data.Entity.Core.Common.CommandTrees.ExpressionBuilder
Imports System.Data.SQLite
'Imports System.Management
'Imports checkid
Public Class frm_main
#Region "======== MAIN_CODE==========================================="
    Dim myversion As String = "38_DEMO"
    Private DBCommand As String = ""
    Private bindingsrc As BindingSource
    Private connstring As String = "Data Source=booklib" & myversion & ".db;Version=3;"
    Private connection As New SQLiteConnection(connstring)
    Private command As New SQLiteCommand("", connection)
    Dim catselect As Integer = 1
    Dim i As Integer




    Private Sub CheckIfDatbaseExist()
        'If System.IO.File.Exists(Application.StartupPath & "booklib" & myversion & ".db") Then
        '    'msgboxX("تماااااااااااااام ")
        'Else
        '    'msgboxX("ملف الداتا غير موجود ")
        '    Me.Close()
        '    Application.Exit()
        'End If
        '' checkid.checkId("26001180201615")
    End Sub
    Public Function IdToBdate(ByVal id As String) As String
        Dim bYear = Mid(id, 2, 2)
        Dim bMonth = Mid(id, 4, 2)
        Dim bDay = Mid(id, 6, 2)
        Dim cent = Mid(id, 1, 1) * 100 + 1700
        Dim bdate = (bDay.ToString) & " / " & (bMonth.ToString) & " / " & ((cent + bYear).ToString)
        Return bdate
    End Function
    Private Sub msgboxX(ByVal msg As String)
        Form1.Label1.Text = msg
        Form1.Show()
    End Sub
    Private Sub frm_main_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ''***********************************
        'Dim g As Graphics = Me.CreateGraphics()
        'MsgBox("ScreenWidth:" & Screen.PrimaryScreen.Bounds.Width & " ScreenHeight:" & Screen.PrimaryScreen.Bounds.Height & vbCrLf & " DpiX:" & g.DpiX & " DpiY:" & g.DpiY)
        '***********************************
        'CheckIfDatbaseExist()
        GroupBox3.Location = New Point(1, 1)
        TabControl1.ItemSize = New Size(0, 1)
        TabControl1.SizeMode = TabSizeMode.Fixed

        Me.Text = Me.Text & "(نظام ادارة المكتبات ( نسخة تجريبية) " & myversion
        Label42.Text = " نظام ادارة المكتبات ( نسخة تجريبية)" & myversion
        showbooks("1")
        'Application.DoEvents()
        fillcat1()
        'Application.DoEvents()
        Fillpublisher()
        'Application.DoEvents()
        showpersons()
        'Application.DoEvents()
        show_borrowed_books()
        '====================================================
        Dim x As Process = Process.GetCurrentProcess()
        Form1.Hide()
        DateTimePicker1.Format = DateTimePickerFormat.Custom
        DateTimePicker1.CustomFormat = "dd-MM-yyyy"
        DateTimePicker2.Format = DateTimePickerFormat.Custom
        DateTimePicker2.CustomFormat = "dd-MM-yyyy"
        'Dim inf As String
        'inf = "Mem Usage: " & x.WorkingSet / 1024 & " K" & vbCrLf _
        '    & "Paged Memory: " & x.PagedMemorySize / 1024 & " K"
        'MessageBox.Show(inf, "Memory Usage")


    End Sub
    Private Sub Button20_Click_1(sender As Object, e As EventArgs) Handles Button20.Click
        export2csv()
    End Sub
    Private Sub Button21_Click(sender As Object, e As EventArgs) Handles Button21.Click
        TabControl1.SelectedTab = TabPage1
    End Sub

    Private Sub Button22_Click(sender As Object, e As EventArgs) Handles Button22.Click
        TabControl1.SelectedTab = TabPage2
    End Sub
    Private Sub Button23_Click(sender As Object, e As EventArgs) Handles Button23.Click
        TabControl1.SelectedTab = TabPage3
    End Sub
    Private Sub Button24_Click(sender As Object, e As EventArgs) Handles Button24.Click
        TabControl1.SelectedTab = TabPage4
    End Sub
    Private Sub Button25_Click(sender As Object, e As EventArgs) Handles Button25.Click
        ' زائر
        GroupBox3.Visible = False
        TabControl1.TabPages.Remove(TabPage6)
        TabControl1.TabPages.Remove(TabPage2)
        TabControl1.TabPages.Remove(TabPage3)
        Button22.Visible = False
        Button23.Visible = False
        Button20.Visible = False
        Button4.Visible = False
        Button5.Visible = False
        Button6.Visible = False
    End Sub

    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
        ' ادارى لادخال كلمة سر
        GroupBox4.Visible = True
        Button25.Visible = False
    End Sub

    Private Sub Button26_Click(sender As Object, e As EventArgs) Handles Button26.Click
        ' عودة
        GroupBox4.Visible = False
        Button25.Visible = True
    End Sub
    Private Sub Button27_Click(sender As Object, e As EventArgs) Handles Button27.Click
        'ادارى
        If TextBox27.Text = "1234" Then
            Button22.Visible = True
            Button23.Visible = True
            GroupBox3.Visible = False
            TabControl1.TabPages.Remove(TabPage6)
        Else
            msgboxX("كلمة السر خطأ")
        End If
    End Sub
#End Region
#Region "=================BOOKS ================================"
    Private Sub searchdata()
        Label3.ForeColor = Color.Red
        Label3.Text = " .... loading ....."
        Cursor = Cursors.WaitCursor
        Application.DoEvents()
        DataGridView1.Rows.Clear()
        connection.Open()
        If connection.State = ConnectionState.Open Then
            command.Connection = connection

            command.CommandText = " SELECT * from books 
                                                        WHERE title Like '%" & TextBox2.Text & "%' 
                                                                And writer Like '%" & TextBox3.Text & "%'
                                                                And bookno like '%" & TextBox17.Text & "%'
                                                                And cat1 Like '%" & catTextBox1.Text & "%'
                                                                And cat2 Like '%" & catTextBox2.Text & "%'
                                                                And cat3 Like '%" & catTextBox3.Text & "%'
                                                                And publisher Like '%" & ComboBox4.Text & "%'
                                                                And bookid like '%" & TextBox7.Text & "%'
                                                        order by FORMAT('%06d', bookno)
                                                        "

            Dim reader As SQLiteDataReader = command.ExecuteReader
            Using reader
                While reader.Read
                    Me.DataGridView1.Rows.Add(reader.GetInt16(0), reader.GetString(1), reader.GetString(2),
                    reader.GetString(3), reader.GetString(4), reader.GetString(5), reader.GetString(6),
                    reader.GetString(7), reader.GetString(8), reader.GetString(9), reader.GetString(10),
                    reader.GetString(11))
                End While
            End Using
        End If
        connection.Close()
        Label3.ForeColor = Color.Black
        Label3.Text = "   عدد الكتب  " & DataGridView1.Rows.Count
        Cursor = Cursors.Default
    End Sub
    Private Sub showbooks(kind As String)
        Dim sortcase As String = " SELECT * from books  order by  bookid LIMIT 20"
        Select Case kind
            Case "1"
                sortcase = " SELECT *  from books   order by  bookid LIMIT 20 "
            Case "2"
                sortcase = " SELECT * from books  order by  trim(title) LIMIT 20"
        End Select
        Label3.ForeColor = Color.Red
        Label3.Text = " .... جارى التحميل ....."
        'Cursor = Cursors.WaitCursor
        'msgboxX("      جارى التحميل     ")

        DataGridView1.Rows.Clear()
        '----FILL TABLE--------------
        connection.Open()
        If connection.State = ConnectionState.Open Then
            command.Connection = connection
            command.CommandText = sortcase
            Dim reader As SQLiteDataReader = command.ExecuteReader
            Using reader
                While reader.Read
                    Me.DataGridView1.Rows.Add(reader.GetInt16(0), reader.GetString(1), reader.GetString(2),
                        reader.GetString(3), reader.GetString(4), reader.GetString(5), reader.GetString(6),
                        reader.GetString(7), reader.GetString(8), reader.GetString(9), reader.GetString(10),
                        reader.GetString(11))
                    Application.DoEvents()
                End While
            End Using
        End If
        connection.Close()
        Label3.ForeColor = Color.Black
        Label3.Text = "   عدد الكتب  " & DataGridView1.Rows.Count
        'Form1.Hide()
        'Cursor = Cursors.Default
    End Sub
    Private Sub clearfields()
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        TextBox5.Text = ""
        TextBox6.Text = ""
        TextBox7.Text = ""
        TextBox10.Text = ""
        TextBox17.Text = ""
        catTextBox1.Text = ""
        catTextBox2.Text = ""
        catTextBox3.Text = ""
        ComboBox1.Text = ""
        ComboBox2.Text = ""
        ComboBox3.Text = ""
        ComboBox4.Text = ""
    End Sub
    Private Sub addbook()
        If TextBox7.Text = "" Or TextBox2.Text = "" Then
            msgboxX("لا يمكن ادخال بيانات كتاب بدون رقم او عنوان")
        Else
            Try
                connection.Open()
                If connection.State = ConnectionState.Open Then
                    command.Connection = connection
                    command.CommandText = "insert into books (bookid, title ,writer,publisher,bookno ,
                                                                                             cat1,cat2,cat3,cab,shelf,publishinfo,notes) 
                                                              values ( '" & TextBox7.Text & " ',
                                                                          '" & TextBox2.Text & " ',
                                                                          '" & TextBox3.Text & " ',
                                                                          '" & ComboBox4.Text & " ' ,
                                                                          '" & TextBox17.Text & " ',
                                                                          '" & catTextBox1.Text & " ',
                                                                          '" & catTextBox2.Text & " ',
                                                                          '" & catTextBox3.Text & " ',
                                                                          '" & TextBox4.Text & " ',
                                                                          '" & TextBox5.Text & " ',
                                                                          '" & TextBox6.Text & " ',
                                                                          '" & TextBox10.Text & " ')"
                    command.ExecuteNonQuery()
                End If
                connection.Close()
            Catch ex As Exception
                If connection.State = ConnectionState.Open Then connection.Close()
                MsgBox(ex.Message)
                'MsgBox("! خطأ رقم الكتاب موجود لكتاب اخر ")
            End Try
            showbooks("1")
        End If
    End Sub
    Private Sub delete_book()
        Dim ask As MsgBoxResult = MsgBox(" هل تريد حذف كتاب رقم  " & TextBox7.Text & " - " & TextBox2.Text, MsgBoxStyle.YesNo)
        If ask = MsgBoxResult.Yes Then
            connection.Open()
            If connection.State = ConnectionState.Open Then
                command.Connection = connection
                command.CommandText = "DELETE from books where bookid ='" & TextBox7.Text & "'"
                command.ExecuteNonQuery()

            End If
            connection.Close()
        End If

        showbooks("1")
    End Sub
    Private Sub update_book()
        connection.Open()
        If connection.State = ConnectionState.Open Then
            command.Connection = connection
            command.CommandText = " update books 
                                                         set  title='" & TextBox2.Text & "',
                                                         writer='" & TextBox3.Text & "',
                                                         cab='" & TextBox4.Text & "',
                                                         shelf='" & TextBox5.Text & "',
                                                         bookno='" & TextBox17.Text & "',
                                                         bookid='" & TextBox7.Text & "',
                                                         cat1='" & catTextBox1.Text & "',
                                                         cat2='" & catTextBox2.Text & "',
                                                         cat3='" & catTextBox3.Text & "',
                                                         publishinfo='" & TextBox6.Text & "',
                                                         publisher='" & ComboBox4.Text & "',
                                                         notes=' " & TextBox10.Text & " '

                                                        where bookid ='" & TextBox7.Text & "'"
            command.ExecuteNonQuery()
        End If
        connection.Close()
        showbooks("1")
    End Sub
    Private Sub Fillpublisher()
        ComboBox4.Items.Clear()
        connection.Open()
        If connection.State = ConnectionState.Open Then
            command.Connection = connection
            command.CommandText = "Select distinct publisher From books order by publisher"
            Dim reader As SQLiteDataReader = command.ExecuteReader
            Using reader
                While reader.Read
                    ComboBox4.Items.Add(reader.GetString(0))
                End While
            End Using
        End If
        connection.Close()
    End Sub
    Private Sub fillcat1()

        ComboBox1.Items.Clear()
        connection.Open()
        If connection.State = ConnectionState.Open Then

            command.Connection = connection
            command.CommandText = "Select * From cat Where catid Like '%00'"

            Dim reader As SQLiteDataReader = command.ExecuteReader

            Using reader
                While reader.Read
                    ComboBox1.Items.Add(reader.GetString(0) & " - " & reader.GetString(1))
                End While
            End Using
        End If
        connection.Close()

        ComboBox2.Items.Clear()
        ComboBox3.Items.Clear()
    End Sub
    Private Sub fillcat2()

        ComboBox2.Items.Clear()
        connection.Open()
        If connection.State = ConnectionState.Open Then
            command.Connection = connection
            command.CommandText = "Select * 
                                                       From cat 
                                                       Where catid Like '" & ComboBox1.SelectedIndex & "%0' 
                                                       AND catid NOT Like '" & ComboBox1.SelectedIndex & "%00' "
            '===================================================================================
            Dim reader As SQLiteDataReader = command.ExecuteReader
            Using reader
                While reader.Read
                    ComboBox2.Items.Add(reader.GetString(0) & " - " & reader.GetString(1))
                End While
            End Using
        End If
        connection.Close()


        ComboBox3.Items.Clear()
    End Sub
    Private Sub fillcat3()
        ComboBox3.Items.Clear()

        connection.Open()
        If connection.State = ConnectionState.Open Then
            command.Connection = connection
            command.CommandText = "Select * 
                                                       From cat  
                                                       Where catid Like '" & ComboBox1.SelectedIndex & ComboBox2.SelectedIndex + 1 & "%'
                                                       AND catid NOT Like '%0' "
            '===================================================================================
            Dim reader As SQLiteDataReader = command.ExecuteReader
            Using reader
                While reader.Read
                    ComboBox3.Items.Add(reader.GetString(0) & " - " & reader.GetString(1))
                End While
            End Using
        End If
        connection.Close()
    End Sub
    Private Sub export2csv()
        Dim sortcase As String = " SELECT * from books  order by  bookno"

        Dim file As New System.IO.StreamWriter("exported_books.csv", True)
        file.WriteLine("رقم الكتاب,العنوان,المؤلف,الناشر,مسلسل,التصنيف,الدولاب,الرف,معلومات  الناشر,ملاحظات")
        Cursor = Cursors.WaitCursor
        msgboxX("      جارى التصدير     ")
        'Application.DoEvents()

        connection.Open()
        If connection.State = ConnectionState.Open Then
            command.Connection = connection
            command.CommandText = sortcase
            Dim reader As SQLiteDataReader = command.ExecuteReader
            Using reader
                While reader.Read
                    file.WriteLine(reader.GetString(4) & "," & reader.GetString(1) & "," & reader.GetString(2) & "," &
                        reader.GetString(3) & "," & reader.GetString(1) & "," & reader.GetString(5) & "," &
                        reader.GetString(8) & "," & reader.GetString(9) & "," & reader.GetString(10) & "," & reader.GetString(11))
                End While
            End Using
        End If
        connection.Close()
        Form1.Hide()
        Cursor = Cursors.Default
        file.Close()
    End Sub
    Private Sub Button1_Click_1(sender As Object, e As EventArgs) Handles Button1.Click
        clearfields()
    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        searchdata()
    End Sub
    Private Sub DataGridView1_MouseDoubleClick(sender As Object, e As MouseEventArgs) Handles DataGridView1.MouseDoubleClick
        Dim irowindex As Integer
        For i As Integer = 0 To DataGridView1.SelectedCells.Count - 1
            irowindex = DataGridView1.SelectedCells.Item(i).RowIndex
            TextBox7.Text = DataGridView1.Rows(irowindex).Cells(0).Value
            TextBox2.Text = DataGridView1.Rows(irowindex).Cells(1).Value
            TextBox3.Text = DataGridView1.Rows(irowindex).Cells(2).Value
            ComboBox4.Text = DataGridView1.Rows(irowindex).Cells(3).Value
            TextBox17.Text = DataGridView1.Rows(irowindex).Cells(4).Value

            catTextBox1.Text = DataGridView1.Rows(irowindex).Cells(5).Value
            catTextBox2.Text = DataGridView1.Rows(irowindex).Cells(6).Value
            catTextBox3.Text = DataGridView1.Rows(irowindex).Cells(7).Value

            If Len(Trim(catTextBox1.Text)) = 3 Then

                ComboBox1.SelectedIndex = CInt(Mid(catTextBox1.Text, 1, 1))
                ComboBox2.SelectedIndex = CInt(Mid(catTextBox1.Text, 2, 1)) - 1
                ComboBox3.SelectedIndex = CInt(Mid(catTextBox1.Text, 3, 1)) - 1
            Else

                ComboBox1.SelectedIndex = -1
                ComboBox2.SelectedIndex = -1
                ComboBox3.SelectedIndex = -1
            End If

            TextBox4.Text = DataGridView1.Rows(irowindex).Cells(8).Value
            TextBox5.Text = DataGridView1.Rows(irowindex).Cells(9).Value
            TextBox6.Text = DataGridView1.Rows(irowindex).Cells(10).Value
            TextBox10.Text = DataGridView1.Rows(irowindex).Cells(11).Value
        Next
    End Sub
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        addbook()
    End Sub
    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles ButtonShowBook2.Click
        showbooks("2")
    End Sub
    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        delete_book()
    End Sub
    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        update_book()
    End Sub
    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        ComboBox3.Text = ""
        fillcat3()
    End Sub
    Private Sub ComboBox3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox3.SelectedIndexChanged
        Select Case catselect
            Case 1
                catTextBox1.Text = ComboBox1.SelectedIndex & ComboBox2.SelectedIndex + 1 & ComboBox3.SelectedIndex + 1
            Case 2
                catTextBox2.Text = ComboBox1.SelectedIndex & ComboBox2.SelectedIndex + 1 & ComboBox3.SelectedIndex + 1
            Case 3
                catTextBox3.Text = ComboBox1.SelectedIndex & ComboBox2.SelectedIndex + 1 & ComboBox3.SelectedIndex + 1
        End Select

    End Sub
    Private Sub catTextBox1_Enter(sender As Object, e As EventArgs) Handles catTextBox1.Enter
        catTextBox1.BackColor = Color.White
        catselect = 1
    End Sub
    Private Sub catTextBox1_Leave(sender As Object, e As EventArgs) Handles catTextBox1.Leave
        catTextBox1.BackColor = Color.FromArgb(255, 255, 128)
    End Sub
    Private Sub catTextBox2_Enter(sender As Object, e As EventArgs) Handles catTextBox2.Enter
        catTextBox2.BackColor = Color.White
        catselect = 2
    End Sub
    Private Sub catTextBox2_Leave(sender As Object, e As EventArgs) Handles catTextBox2.Leave
        catTextBox2.BackColor = Color.FromArgb(255, 255, 128)
    End Sub
    Private Sub catTextBox3_Enter(sender As Object, e As EventArgs) Handles catTextBox3.Enter
        catTextBox3.BackColor = Color.White
        catselect = 3
    End Sub
    Private Sub catTextBox3_Leave(sender As Object, e As EventArgs) Handles catTextBox3.Leave
        catTextBox3.BackColor = Color.FromArgb(255, 255, 128)
    End Sub
    Private Sub ComboBox1_SelectedIndexChanged_1(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        ComboBox2.Text = ""
        ComboBox3.Text = ""
        fillcat2()
    End Sub
#End Region
#Region "===================== PERSONS ==================================="
    Private Sub showpersons()
        Label29.ForeColor = Color.Red
        Label29.Text = " .... loading ....."
        Cursor = Cursors.WaitCursor
        'Application.DoEvents()

        DataGridView2.Rows.Clear()
        connection.Open()
        If connection.State = ConnectionState.Open Then
            command.Connection = connection
            command.CommandText = " SELECT * from persons  order by  id "
            Dim reader As SQLiteDataReader = command.ExecuteReader
            Using reader
                While reader.Read
                    Me.DataGridView2.Rows.Add(reader.GetInt16(0), reader.GetString(1), reader.GetString(2),
                                                                      reader.GetString(3), reader.GetString(4), reader.GetString(5),
                                                                      reader.GetString(6), reader.GetString(7))
                    Application.DoEvents()

                End While
            End Using
        End If
        connection.Close()
        Label29.ForeColor = Color.Black
        Label29.Text = "   عدد المشتركين  " & DataGridView2.Rows.Count


        'For i As Integer = 0 To Me.DataGridView2.Rows.Count - 1
        '    If chkId(Me.DataGridView2.Rows(i).Cells(2)) Then

        '        'DataGridView2.Rows(i).Cells(2).Style.BackColor = Color.Red
        '        DataGridView2.Rows(i).DefaultCellStyle.BackColor = Color.Red
        '    End If
        'Next

        For i As Integer = 0 To Me.DataGridView2.Rows.Count - 1
            If (Mid(Me.DataGridView2.Rows(i).Cells(2).Value, 1, 1) = "0") Then
                'DataGridView2.Rows(i).Cells(2).Style.BackColor = Color.Red
                DataGridView2.Rows(i).DefaultCellStyle.BackColor = Color.Red
            End If
        Next

        Cursor = Cursors.Default
    End Sub
    Private Sub addperson()
        If TextBox1.Text = "" Or TextBox9.Text = "" Or TextBox14.Text = "" Then
            msgboxX("لا يمكن ادخال بيانات مشترك بدون رقم او أسم او رقم قومى")
        Else
            Try
                connection.Open()
                If connection.State = ConnectionState.Open Then
                    command.Connection = connection
                    command.CommandText = "insert into persons (id,name,natid,tel,grade,study,village,notes)
                                                              values ( '" & TextBox1.Text & " ',
                                                                          '" & TextBox9.Text & " ',
                                                                          '" & TextBox14.Text & " ',
                                                                          '" & TextBox15.Text & " ',
                                                                          '" & TextBox16.Text & " ',
                                                                         '" & TextBox18.Text & " ',
                                                                          '" & TextBox20.Text & " ',
                                                                          '" & TextBox21.Text & "')"
                    command.ExecuteNonQuery()
                End If
                connection.Close()
            Catch ex As Exception
                If connection.State = ConnectionState.Open Then connection.Close()
                MsgBox(ex.Message)

            End Try
            showpersons()
        End If
    End Sub
    Private Sub delete_person()
        Dim ask As MsgBoxResult = MsgBox(" هل تريد حذف مشترك  رقم  " & TextBox1.Text & " - " & TextBox9.Text, MsgBoxStyle.YesNo)
        If ask = MsgBoxResult.Yes Then
            connection.Open()
            If connection.State = ConnectionState.Open Then
                command.Connection = connection
                command.CommandText = "DELETE from persons where id='" & TextBox1.Text & "'"
                command.ExecuteNonQuery()
            End If
            connection.Close()

        End If
        showpersons()
    End Sub
    Private Sub clearPersonsField()

        Dim aaa As Control
        For Each aaa In GroupBox2.Controls
            If aaa.Tag = "ppp" Then
                aaa.Text = ""
            End If
        Next
    End Sub
    Private Sub searchpersons()
        Label29.ForeColor = Color.Red
        Label29.Text = " .... loading ....."
        Cursor = Cursors.WaitCursor
        'Application.DoEvents()

        DataGridView2.Rows.Clear()
        connection.Open()
        If connection.State = ConnectionState.Open Then
            command.Connection = connection
            command.CommandText = " SELECT * from persons 
                                                        WHERE id like '%" & TextBox1.Text & "%' 
                                                        And name Like '%" & TextBox9.Text & "%'
                                                        And natid like '%" & TextBox14.Text & "%'
                                                        And tel Like '%" & TextBox15.Text & "%'
                                                        order by  id"

            Dim reader As SQLiteDataReader = command.ExecuteReader
            Using reader
                While reader.Read
                    Me.DataGridView2.Rows.Add(reader.GetInt16(0), reader.GetString(1), reader.GetString(2),
                    reader.GetString(3), reader.GetString(4), reader.GetString(5), reader.GetString(6), reader.GetString(7))
                End While
            End Using
        End If
        connection.Close()
        Label29.ForeColor = Color.Black
        Label29.Text = "   عدد المشتركين :" & DataGridView2.Rows.Count
        Cursor = Cursors.Default
    End Sub
    Private Sub UpdatePerson()
        connection.Open()
        If connection.State = ConnectionState.Open Then
            command.Connection = connection
            command.CommandText = " update persons
                                                         set 
                                                               id='" & TextBox1.Text & "',
                                                         name='" & TextBox9.Text & "',
                                                          natid='" & TextBox14.Text & "',
                                                              tel='" & TextBox15.Text & "',
                                                         grade='" & TextBox16.Text & "',
                                                        village='" & TextBox18.Text & "',
                                                        study='" & TextBox20.Text & "',
                                                          notes='" & TextBox21.Text & "'

                                                        where id ='" & TextBox1.Text & "'"
            command.ExecuteNonQuery()
        End If
        connection.Close()
        showpersons()
    End Sub
    Private Sub DataGridView2_MouseDoubleClick(sender As Object, e As MouseEventArgs) Handles DataGridView2.MouseDoubleClick
        Dim irowindex As Integer
        For i As Integer = 0 To DataGridView2.SelectedCells.Count - 1
            irowindex = DataGridView2.SelectedCells.Item(i).RowIndex
            TextBox1.Text = Trim(DataGridView2.Rows(irowindex).Cells(0).Value)
            TextBox9.Text = DataGridView2.Rows(irowindex).Cells(1).Value
            TextBox14.Text = DataGridView2.Rows(irowindex).Cells(2).Value
            TextBox15.Text = DataGridView2.Rows(irowindex).Cells(3).Value
            TextBox16.Text = DataGridView2.Rows(irowindex).Cells(4).Value
            TextBox18.Text = DataGridView2.Rows(irowindex).Cells(5).Value
            TextBox20.Text = DataGridView2.Rows(irowindex).Cells(6).Value
            TextBox21.Text = DataGridView2.Rows(irowindex).Cells(7).Value
            Label28.Text = (Date.Now.Year) -
                ((Mid(DataGridView2.Rows(irowindex).Cells(2).Value, 1, 1) * 100) + 1700 +
                (Mid(DataGridView2.Rows(irowindex).Cells(2).Value, 2, 2)))
            Label31.Text = IdToBdate(TextBox14.Text)
            If (Label28.Text) > 100 Or (Label28.Text) < 4 Then
                Label28.BackColor = Color.Red
            Else
                Label28.BackColor = Color.White '
            End If
        Next
    End Sub
    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        addperson()
    End Sub
    Private Sub member_Button1_Click(sender As Object, e As EventArgs) Handles member_Button1.Click
        delete_person()
    End Sub
    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        clearPersonsField()
    End Sub
    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        searchpersons()
    End Sub
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        UpdatePerson()
    End Sub
    Private Sub Button12_Click(sender As Object, e As EventArgs) Handles Button12.Click
        showpersons()
    End Sub
#End Region
#Region "============ BORROW ======================================="
    Private Sub show_borrowed_books()
        Cursor = Cursors.WaitCursor
        DataGridView3.Rows.Clear()
        connection.Open()
        If connection.State = ConnectionState.Open Then
            command.Connection = connection
            command.CommandText = "SELECT borrowid,
                                                                            w.bookid, 
                                                                            b.title , 
                                                                            w.personid, 
                                                                            p.name,
                                                                          w.borrowdate,
                                                                          w.period,
                                                                          date(w.borrowdate, '+' || w.period || ' day'),
                                                                          w.actualretdate,
                                               JULIANDAY(w.actualretdate) - JULIANDAY(w.borrowdate) as sd
                                                                         
                                                                        
                                                           FROM persons p 
                                                            JOIN books b 
                                                            JOIN borrows w 
                                                            WHERE w.personId = p.id and w.bookid = b.bookid
                                                                order by sd desc"
            'JULIANDAY(w.actualretdate) - JULIANDAY(w.borrowdate)
            Dim reader As SQLiteDataReader = command.ExecuteReader
            Using reader
                While reader.Read
                    Me.DataGridView3.Rows.Add(reader.GetInt16(0),
                                                                       reader.GetInt16(1),
                                                                       reader.GetString(2),
                                                                              reader.GetInt16(3),
                                                                            reader.GetString(4),
                                                                                         reader.GetString(5),
                                                                                            reader.GetInt16(6),
                                                                                            reader.GetString(7),
                                                                                            reader.GetString(8),
                                                                                            reader.GetDouble(9))

                    Application.DoEvents()
                End While
            End Using
        End If
        connection.Close()

        'For i As Integer = 0 To Me.DataGridView2.Rows.Count - 1
        '    If chkId(Me.DataGridView2.Rows(i).Cells(2)) Then

        '        'DataGridView2.Rows(i).Cells(2).Style.BackColor = Color.Red
        '        DataGridView2.Rows(i).DefaultCellStyle.BackColor = Color.Red
        '    End If
        'Next
        For i As Integer = 0 To DataGridView3.Rows.Count - 1
            If ((DataGridView3.Rows(i).Cells(9).Value) > (DataGridView3.Rows(i).Cells(6).Value)) Then
                DataGridView3.Rows(i).DefaultCellStyle.BackColor = Color.Red
            End If
        Next

        Cursor = Cursors.Default
    End Sub
    Private Sub clearBorrow()
        Dim aaa As Control
        For Each aaa In GroupBox1.Controls
            If aaa.Tag = "bbb" Then
                aaa.Text = ""
            End If
        Next

    End Sub
    Private Sub addBorrow()
        clearBorrow()
        TextBox13.Text = DataGridView3.Rows.Count + 1
        TextBox12.Text = TextBox7.Text
        TextBox11.Text = TextBox2.Text
        TextBox19.Text = TextBox1.Text
        TextBox8.Text = TextBox9.Text
        TextBox24.Text = 10

    End Sub
    Private Sub saveborrow()

        If TextBox13.Text = "" Or TextBox12.Text = "" Or TextBox11.Text = "" Then
            msgboxX("لا يمكن ادخال استعارة بدون رقم او كتاب او مشترك")
        Else
            Try
                connection.Open()
                If connection.State = ConnectionState.Open Then
                    command.Connection = connection
                    command.CommandText = "insert into borrows (borrowid, bookid,
                                                                  personid,borrowdate,period,
                                                                  actualretdate,notes) 
                                                              values ( '" & TextBox13.Text & " ',
                                                                          '" & TextBox12.Text & " ',
                                                                          '" & TextBox19.Text & " ',
                                                                          '" & DateTimePicker1.Value.Date.ToString("yyyy-MM-dd") & " ',
                                                                          '" & TextBox24.Text & " ',
                                                                          '" & DateTimePicker2.Value.Date.ToString("yyyy-MM-dd") & " ',
                                                                          '" & TextBox26.Text & " ')"
                    command.ExecuteNonQuery()
                End If
                connection.Close()
            Catch ex As Exception
                If connection.State = ConnectionState.Open Then connection.Close()
                MsgBox(ex.Message)
                'MsgBox("! خطأ رقم الكتاب موجود لكتاب اخر ")
            End Try
            'showborrow()
        End If
    End Sub
    Private Sub DataGridView3_MouseDoubleClick(sender As Object, e As MouseEventArgs) Handles DataGridView3.MouseDoubleClick
        Dim irowindex As Integer
        For i As Integer = 0 To DataGridView3.SelectedCells.Count - 1
            irowindex = DataGridView3.SelectedCells.Item(i).RowIndex
            TextBox13.Text = DataGridView3.Rows(irowindex).Cells(0).Value
            TextBox12.Text = DataGridView3.Rows(irowindex).Cells(1).Value
            TextBox11.Text = DataGridView3.Rows(irowindex).Cells(2).Value
            TextBox19.Text = DataGridView3.Rows(irowindex).Cells(3).Value
            TextBox8.Text = DataGridView3.Rows(irowindex).Cells(4).Value

            TextBox24.Text = DataGridView3.Rows(irowindex).Cells(6).Value
            TextBox23.Text = DataGridView3.Rows(irowindex).Cells(7).Value

            TextBox26.Text = DataGridView3.Rows(irowindex).Cells(9).Value

            DateTimePicker1.Value = Convert.ToDateTime(DataGridView3.Rows(irowindex).Cells(5).Value)
            DateTimePicker2.Value = Convert.ToDateTime(DataGridView3.Rows(irowindex).Cells(8).Value)
        Next
    End Sub
    Private Sub delete_borrow()
        Dim ask As MsgBoxResult = MsgBox(" هل تريد حذف استعارة رقم  " & TextBox13.Text, MsgBoxStyle.YesNo)
        If ask = MsgBoxResult.Yes Then
            connection.Open()
            If connection.State = ConnectionState.Open Then
                command.Connection = connection
                command.CommandText = "DELETE from borrows where borrowid ='" & TextBox13.Text & "'"
                command.ExecuteNonQuery()

            End If
            connection.Close()
        End If

        showbooks("1")
    End Sub
    Private Sub Button16_Click(sender As Object, e As EventArgs) Handles Button16.Click
        addBorrow()
    End Sub
    Private Sub Button19_Click(sender As Object, e As EventArgs) Handles Button19.Click
        clearBorrow()
    End Sub
    Private Sub Button15_Click(sender As Object, e As EventArgs) Handles Button15.Click
        saveborrow()
        show_borrowed_books()

    End Sub
    Private Sub DateTimePicker1_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker1.ValueChanged
        Dim asd As Date = DateAdd("d", Convert.ToInt32(TextBox24.Text), DateTimePicker1.Value)
        TextBox23.Text = Format(asd, "dd, MMM, yyyy")
    End Sub

    Private Sub Button17_Click(sender As Object, e As EventArgs) Handles Button17.Click
        show_borrowed_books()
    End Sub

    Private Sub Button18_Click(sender As Object, e As EventArgs) Handles Button18.Click
        delete_borrow()
        show_borrowed_books()
    End Sub




#End Region
End Class