Imports Microsoft.Office.Interop


Public Class Form2
    Dim objExcel As New Excel.Application     ' Represents an instance of Excel
    Dim objWorkbook As Excel.Workbook     'Represents a workbook object
    Dim objWorksheet As Excel.Worksheet     'Represents a worksheet object

   
    Private Sub SaveFileDialog1_FileOk(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs)

    End Sub

    Private Sub TreeView1_AfterSelect(ByVal sender As System.Object, ByVal e As System.Windows.Forms.TreeViewEventArgs)

    End Sub

    Private Sub TreeView1_BeforeCollapse(ByVal sender As Object, ByVal e As System.Windows.Forms.TreeViewCancelEventArgs)
       

    End Sub

    Private Sub Form2_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'Dim objExcel As New Excel.Application     ' Represents an instance of Excel
        'Dim objWorkbook As Excel.Workbook     'Represents a workbook object
        'Dim objWorksheet As Excel.Worksheet     'Represents a worksheet object

        'Add a new workbook
        'objWorkbook = objExcel.Workbooks.Add

        'NB: If you wanted to open an existing workbook then the following could replace 
        'the above line

        objWorkbook = objExcel.Workbooks.Open(Form3.TextBox1.Text)

        'Set the WorkSheet object to the sheet in the workbook you want to use
        'NB: You can use an Index number as well as specifying the name of the sheet
        objWorksheet = CType(objWorkbook.Worksheets.Item("Sheet1"), Excel.Worksheet)

        'This form contains two text boxes, write values to cells A1 and A2
        ' objWorksheet.Cells(1, 1) = TextBox1.Text
        'objWorksheet.Cells(2, 1) = TextBox2.Text


        'Read data from cells C18 and F18
        TextBox1.Text = objWorksheet.Cells(18, 3).value
        TextBox2.Text = objWorksheet.Cells(18, 5).value
        TextBox3.Text = objWorksheet.Cells(8, 1).value
        TextBox4.Text = objWorksheet.Cells(9, 1).value
        TextBox5.Text = objWorksheet.Cells(1, 20).value
        TextBox6.Text = objWorksheet.Cells(19, 2).value
        TextBox7.Text = objWorksheet.Cells(18, 2).value
        TextBox8.Text = objWorksheet.Cells(21, 2).value
        TextBox9.Text = objWorksheet.Cells(22, 2).value
        TextBox10.Text = objWorksheet.Cells(21, 3).value
        TextBox11.Text = objWorksheet.Cells(21, 5).value
        TextBox12.Text = objWorksheet.Cells(2, 20).value
        TextBox13.Text = objWorksheet.Cells(24, 2).value
        TextBox14.Text = objWorksheet.Cells(25, 2).value
        TextBox15.Text = objWorksheet.Cells(24, 3).value
        TextBox16.Text = objWorksheet.Cells(24, 5).value
        TextBox17.Text = objWorksheet.Cells(3, 20).value
        TextBox18.Text = objWorksheet.Cells(27, 2).value
        TextBox19.Text = objWorksheet.Cells(28, 2).value
        TextBox20.Text = objWorksheet.Cells(27, 3).value
        TextBox21.Text = objWorksheet.Cells(27, 5).value
        TextBox22.Text = objWorksheet.Cells(4, 20).value
        TextBox23.Text = objWorksheet.Cells(30, 2).value
        TextBox24.Text = objWorksheet.Cells(31, 2).value
        TextBox25.Text = objWorksheet.Cells(30, 3).value
        TextBox26.Text = objWorksheet.Cells(30, 5).value
        TextBox27.Text = objWorksheet.Cells(5, 20).value
        TextBox28.Text = objWorksheet.Cells(33, 2).value
        TextBox29.Text = objWorksheet.Cells(34, 2).value
        TextBox30.Text = objWorksheet.Cells(33, 3).value
        TextBox31.Text = objWorksheet.Cells(33, 5).value
        TextBox32.Text = objWorksheet.Cells(6, 20).value
        TextBox33.Text = objWorksheet.Cells(10, 1).value
        
        'Close the workbook and Excel
        ' objWorkbook.Close(False)

        'NB: Above will not save the changes, to save a workbook before closing 
        'use the Save or SaveAs method of the workbook object:

        'objWorkbook.Save()
        'or
        'objWorkbook.SaveAs("C:\Temp\Book1.xls")

        'objExcel.Quit()
    End Sub

    Private Sub TextBox2_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox2.TextChanged

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        objWorksheet.Cells(18, 3).value = TextBox1.Text
        objWorksheet.Cells(18, 5).value = TextBox2.Text
        objWorksheet.Cells(8, 1).value = TextBox3.Text
        objWorksheet.Cells(9, 1).value = TextBox4.Text
        objWorksheet.Cells(1, 20).value = TextBox5.Text
        objWorksheet.Cells(19, 2).value = TextBox6.Text
        objWorksheet.Cells(18, 2).value = TextBox7.Text
        objWorksheet.Cells(21, 2).value = TextBox8.Text
        objWorksheet.Cells(22, 2).value = TextBox9.Text
        objWorksheet.Cells(21, 3).value = TextBox10.Text
        objWorksheet.Cells(21, 5).value = TextBox11.Text
        objWorksheet.Cells(2, 20).value = TextBox12.Text
        objWorksheet.Cells(24, 2).value = TextBox13.Text
        objWorksheet.Cells(25, 2).value = TextBox14.Text
        objWorksheet.Cells(24, 3).value = TextBox15.Text
        objWorksheet.Cells(24, 5).value = TextBox16.Text
        objWorksheet.Cells(3, 20).value = TextBox17.Text
        objWorksheet.Cells(27, 2).value = TextBox18.Text
        objWorksheet.Cells(28, 2).value = TextBox19.Text
        objWorksheet.Cells(27, 3).value = TextBox20.Text
        objWorksheet.Cells(27, 5).value = TextBox21.Text
        objWorksheet.Cells(4, 20).value = TextBox22.Text
        objWorksheet.Cells(30, 2).value = TextBox23.Text
        objWorksheet.Cells(31, 2).value = TextBox24.Text
        objWorksheet.Cells(30, 3).value = TextBox25.Text
        objWorksheet.Cells(30, 5).value = TextBox26.Text
        objWorksheet.Cells(33, 2).value = TextBox28.Text
        objWorksheet.Cells(34, 2).value = TextBox29.Text
        objWorksheet.Cells(33, 3).value = TextBox30.Text
        objWorksheet.Cells(33, 5).value = TextBox31.Text
        objWorksheet.Cells(6, 20).value = TextBox32.Text

        'date and time
        Dim a As String
        Dim b As String
        a = "Date : "
        b = a & DateTimePicker1.Value.ToString("MMM dd, yyyy")

        'perform subtotal
        Dim total As Double
        Dim tax1 As Double
        Dim tax2 As Double
        Dim tax3 As Double
        Dim tax4 As Double
        Dim tax5 As Double
        Dim tax6 As Double

        TextBox1.Text = 0
        TextBox2.Text = 0
        TextBox5.Text = 0
        tax1 = TextBox1.Text * TextBox2.Text * (TextBox5.Text / 100)
        TextBox10.Text = 0
        TextBox12.Text = 0
        TextBox11.Text = 0
        tax2 = TextBox10.Text * TextBox11.Text * (TextBox12.Text / 100)
        TextBox15.Text = 0
        TextBox16.Text = 0
        TextBox17.Text = 0
        tax3 = TextBox15.Text * TextBox16.Text * (TextBox17.Text / 100)
        TextBox20.Text = 0
        TextBox21.Text = 0
        TextBox22.Text = 0
        tax4 = TextBox20.Text * TextBox21.Text * (TextBox22.Text / 100)
        TextBox25.Text = 0
        TextBox26.Text = 0
        TextBox27.Text = 0
        tax5 = TextBox25.Text * TextBox26.Text * (TextBox27.Text / 100)
        TextBox30.Text = 0
        TextBox31.Text = 0
        TextBox32.Text = 0
        tax6 = TextBox30.Text * TextBox31.Text * (TextBox32.Text / 100)


        'perform tax
        total = tax1 + tax2 + tax3 + tax4 + tax5 + tax6

        'perform total
        Dim amount1 As Integer
        Dim amount2 As Integer
        Dim amount3 As Integer
        Dim amount4 As Integer
        Dim amount5 As Integer
        Dim amount6 As Integer
        Dim amount As Integer
        amount1 = TextBox1.Text * TextBox2.Text
        amount2 = TextBox10.Text * TextBox11.Text
        amount3 = TextBox15.Text * TextBox16.Text
        amount4 = TextBox20.Text * TextBox21.Text
        amount5 = TextBox25.Text * TextBox26.Text
        amount6 = TextBox30.Text * TextBox31.Text
        amount = amount1 + amount2 + amount3 + amount4 + amount5 + amount6

        With objWorksheet
            .Cells(11, "E").value = b       'date
        End With

        With objWorksheet
            If TextBox7.Text <> "" Then
                .Cells(18, "A").value = "1"
                .Cells(18, "B").value = TextBox7.Text
                .Cells(18, "C").value = TextBox1.Text
                .Cells(18, "E").value = TextBox2.Text
                .Cells(18, "F").value = amount1
                .Cells(1, "T").value = TextBox5.Text
                .Cells(36, "F").value = amount                 'total amount

                .Cells(37, "F").value = total                  ' total taxes

                .Cells(39, "F").value = total + amount          'grand total

                .Cells(19, "B").value = TextBox6.Text
            

            End If

            If TextBox8.Text <> "" Then
                .Cells(21, "A").value = "2"
                .Cells(21, "B").value = TextBox8.Text
                .Cells(21, "C").value = TextBox10.Text
                .Cells(21, "E").value = TextBox11.Text
                .Cells(21, "F").value = amount2
                .Cells(22, "B").value = TextBox9.Text
                .Cells(2, "T").value = TextBox12.Text
         
            End If

            If TextBox13.Text <> "" Then
                .Cells(24, "A").value = "3"
                .Cells(24, "B").value = TextBox13.Text
                .Cells(24, "C").value = TextBox15.Text
                .Cells(24, "E").value = TextBox16.Text
                .Cells(24, "F").value = amount3
                .Cells(25, "B").value = TextBox14.Text
                .Cells(3, "T").value = TextBox17.Text
           
            End If

            If TextBox18.Text <> "" Then
                .Cells(27, "A").value = "4"
                .Cells(27, "B").value = TextBox18.Text
                .Cells(27, "C").value = TextBox20.Text
                .Cells(27, "E").value = TextBox21.Text
                .Cells(27, "F").value = amount4
                .Cells(28, "B").value = TextBox19.Text
                .Cells(4, "T").value = TextBox22.Text
          
            End If

            If TextBox23.Text <> "" Then
                .Cells(30, "A").value = "5"
                .Cells(30, "B").value = TextBox23.Text
                .Cells(30, "C").value = TextBox25.Text
                .Cells(30, "E").value = TextBox26.Text
                .Cells(30, "F").value = amount5
                .Cells(31, "B").value = TextBox24.Text
                .Cells(5, "T").value = TextBox27.Text
         
            End If

            If TextBox28.Text <> "" Then
                .Cells(33, "A").value = "6"
                .Cells(33, "B").value = TextBox28.Text
                .Cells(33, "C").value = TextBox30.Text
                .Cells(33, "E").value = TextBox31.Text
                .Cells(33, "F").value = amount6
                .Cells(34, "B").value = TextBox29.Text
                .Cells(6, "T").value = TextBox32.Text
            
            End If

        End With

        objWorkbook.Save()

        'Close the workbook and Excel
        objWorkbook.Close()

        objExcel.Quit()
    End Sub

    Private Sub EditMoreToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles EditMoreToolStripMenuItem.Click
        Form3.Show()
        Me.Close()
    End Sub

    Private Sub GenerateNewToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles GenerateNewToolStripMenuItem.Click
        Form1.Show()
        Me.Close()
    End Sub

    Private Sub ExitToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExitToolStripMenuItem.Click
        Me.Close()
    End Sub
End Class