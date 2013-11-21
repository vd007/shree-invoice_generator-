Imports Excel = Microsoft.Office.Interop.Excel
Imports System.IO
Imports System.Data.OleDb

Public Class Form1
    '~~> Define your Excel Objects
    Dim xlApp As New Excel.Application
    Dim xlWorkBook As Excel.Workbook
    Dim xlWorkSheet As Excel.Worksheet


    'row operation
    Dim row As Integer = 1
    Dim rowstr As String

    'random number
    Dim randomstr As String

    'column operation
    Dim con As String
    Dim column As Integer = 65
    Dim colvalue As Char

    Dim sw As StreamWriter
    Dim path As String = "C:\shree\files\names.txt"
    Dim pathitems As String = "C:\shree\files\items.txt"
    Dim place As String = "C:\shree\files\places.txt"
    Dim desc As String = "C:\shree\files\desc.txt"

    Dim filename As String


    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Button2.Enabled = True
        Button1.Enabled = False
        Button3.Enabled = True

        randomnumber()

        'write all the data related to user

        'Dim text6 As String
        'Dim text8 As String
        'text6 = TextBox6.Text
        'text8 = TextBox8.Text
        'If text6 <> text8 Then
        '    MsgBox("taxes must be same for all the products")
        '    TextBox6.Text = ""
        '    TextBox8.Text = ""
        'Else
        'initialize new workbook
        startwork()
        'End If





        'call reset
        reset()



        '------------------------------------------------------------------------------'
        'file operation


        ' This text is always added, making the file longer over time 
        ' if it is not deleted. 
        Dim appendText As String = Environment.NewLine + ComboBox1.Text

        Dim appendTextitem1 As String = Environment.NewLine + ComboBox2.Text
        Dim appendTextitem2 As String = Environment.NewLine + ComboBox3.Text
        Dim appendTextitem3 As String = Environment.NewLine + ComboBox4.Text
        Dim appendTextitem4 As String = Environment.NewLine + ComboBox5.Text
        Dim appendTextitem5 As String = Environment.NewLine + ComboBox6.Text
        Dim appendTextitem6 As String = Environment.NewLine + ComboBox7.Text

        Dim appentextitem7 As String = Environment.NewLine + ComboBox9.Text

        Dim appendtext8 As String = Environment.NewLine + ComboBox10.Text
        Dim appendtext9 As String = Environment.NewLine + ComboBox11.Text
        Dim appendtext10 As String = Environment.NewLine + ComboBox12.Text
        Dim appendtext11 As String = Environment.NewLine + ComboBox13.Text
        Dim appendtext12 As String = Environment.NewLine + ComboBox14.Text
        Dim appendtext13 As String = Environment.NewLine + ComboBox15.Text

        Dim comboblank As String = "" + ControlChars.Back  'add a blank character + backspace


        If ComboBox1.Items.Contains(ComboBox1.Text) Then
        Else
            File.AppendAllText(path, appendText)
            'MsgBox(ComboBox1.Items.Count)
        End If
        If ComboBox2.Items.Contains(ComboBox2.Text) Then
        Else
            File.AppendAllText(pathitems, appendTextitem1)

        End If
        If ComboBox3.Items.Contains(ComboBox3.Text) Then
        Else
            File.AppendAllText(pathitems, appendTextitem2)
        End If
        If ComboBox4.Items.Contains(ComboBox4.Text) Then
        Else
            File.AppendAllText(pathitems, appendTextitem3)
        End If
        If ComboBox5.Items.Contains(ComboBox5.Text) Then
        Else
            File.AppendAllText(pathitems, appendTextitem4)
        End If
        If ComboBox6.Items.Contains(ComboBox6.Text) Then
        Else
            File.AppendAllText(pathitems, appendTextitem5)
        End If
        If ComboBox7.Items.Contains(ComboBox7.Text) Then
        Else
            File.AppendAllText(pathitems, appendTextitem6)
        End If
        If ComboBox9.Items.Contains(ComboBox9.Text) Then
        Else
            File.AppendAllText(place, appentextitem7)
            'MsgBox(ComboBox1.Items.Count)
        End If
        If ComboBox10.Items.Contains(ComboBox10.Text) Then
        Else
            File.AppendAllText(desc, appendtext8)

        End If
        If ComboBox11.Items.Contains(ComboBox11.Text) Then
        Else
            File.AppendAllText(desc, appendtext9)

        End If
        If ComboBox12.Items.Contains(ComboBox12.Text) Then
        Else
            File.AppendAllText(desc, appendtext10)

        End If
        If ComboBox13.Items.Contains(ComboBox13.Text) Then
        Else
            File.AppendAllText(desc, appendtext11)

        End If
        If ComboBox14.Items.Contains(ComboBox14.Text) Then
        Else
            File.AppendAllText(desc, appendtext12)

        End If
        If ComboBox15.Items.Contains(ComboBox15.Text) Then
        Else
            File.AppendAllText(desc, appendtext13)

        End If
        'fill the default values
        'If ComboBox1.Text = "" Then
        '    File.AppendAllText(path, comboblank)
        'Else
        '    File.AppendAllText(path, appendText)
        'End If
        'If ComboBox2.Text = "" Then
        '    File.AppendAllText(pathitems, comboblank)
        'Else
        '    File.AppendAllText(pathitems, appendTextitem1)
        'End If
        'If ComboBox3.Text = "" Then
        '    File.AppendAllText(pathitems, comboblank)
        'Else
        '    File.AppendAllText(pathitems, appendTextitem2)
        'End If
        'If ComboBox4.Text = "" Then
        '    File.AppendAllText(pathitems, comboblank)
        'Else
        '    File.AppendAllText(pathitems, appendTextitem3)
        'End If
        'If ComboBox5.Text = "" Then
        '    File.AppendAllText(pathitems, comboblank)
        'Else
        '    File.AppendAllText(pathitems, appendTextitem4)
        'End If
        'If ComboBox6.Text = "" Then
        '    File.AppendAllText(pathitems, comboblank)
        'Else
        '    File.AppendAllText(pathitems, appendTextitem5)
        'End If
        'If ComboBox7.Text = "" Then
        '    File.AppendAllText(pathitems, comboblank)
        'Else
        '    File.AppendAllText(pathitems, appendTextitem6)
        'End If


    End Sub
    Private Sub movecell(ByVal a) 'function increment row number and return the required string
        row = row + a
        rowstr = row
        colvalue = Chr(column)
        con = colvalue + rowstr
    End Sub
    Private Sub reset()
        row = 1
        column = 65
    End Sub
    Private Sub movecolumn(ByVal a) 'function increment column value
        column = column + a
    End Sub
    Private Sub randomnumber()
        Dim random As Integer

        random = CInt(Int((99999999 * Rnd()) + 10000000))
        randomstr = random
    End Sub
    Private Sub startwork()
        '~~> Add a New Workbook
        xlWorkBook = xlApp.Workbooks.Add

        '~~> Display Excel
        xlApp.Visible = True

        '~~> Set the relebant sheet that we want to work with
        xlWorkSheet = xlWorkBook.Sheets("Sheet1")

        'savework()

        'date and time
        Dim a As String
        Dim b As String
        a = "Date : "
        b = a & DateTimePicker1.Value.ToString("MMM dd, yyyy")

        'perform subtotal
        Dim text4 As Integer
        Dim text5 As Integer
        Dim total As Integer
        Dim tax1 As Integer
        Dim text7 As Integer
        Dim text2 As Integer
        Dim tax2 As Integer
        Dim tax3 As Integer
        Dim tax4 As Integer
        Dim tax5 As Integer
        Dim tax6 As Integer
        text2 = 0
        text5 = 0

        text4 = TextBox4.Text
        text5 = TextBox5.Text
        tax1 = text5 * text4 * (TextBox6.Text / 100)
        text2 = 0
        text7 = 0
        text2 = TextBox2.Text
        text7 = TextBox7.Text
        tax2 = text2 * text7 * (TextBox8.Text / 100)
        TextBox12.Text = 0
        TextBox13.Text = 0
        TextBox14.Text = 0
        tax3 = TextBox12.Text * TextBox13.Text * (TextBox14.Text / 100)
        TextBox17.Text = 0
        TextBox18.Text = 0
        TextBox19.Text = 0
        tax4 = TextBox17.Text * TextBox18.Text * (TextBox19.Text / 100)
        TextBox22.Text = 0
        TextBox23.Text = 0
        TextBox24.Text = 0
        tax5 = TextBox22.Text * TextBox23.Text * (TextBox24.Text / 100)
        TextBox27.Text = 0
        TextBox28.Text = 0
        TextBox29.Text = 0
        tax6 = TextBox27.Text * TextBox28.Text * (TextBox29.Text / 100)


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
        amount1 = (text4 * text5)
        amount2 = (TextBox2.Text * TextBox7.Text)
        amount3 = (TextBox12.Text * TextBox13.Text)
        amount4 = (TextBox17.Text * TextBox18.Text)
        amount5 = (TextBox22.Text * TextBox23.Text)
        amount6 = (TextBox27.Text * TextBox28.Text)
        amount = amount1 + amount2 + amount3 + amount4 + amount5 + amount6


        'calculate taxes
        Dim tax As Integer
        tax = (TextBox6.Text * total) / 100

        'generate quatation number


        '=====================================================================
        With xlWorkSheet
            .Cells(11, "E").value = b       'date
            .Cells(9, "A").value = TextBox10.Text   'address
            .Cells(10, "A").value = ComboBox9.Text  'place
        End With
        With xlWorkSheet
            If ComboBox2.Text <> "" Then
                .Cells(18, "A").value = "1"
                .Cells(18, "B").value = ComboBox2.Text
                .Cells(18, "C").value = TextBox4.Text
                .Cells(18, "E").value = TextBox5.Text
                .Cells(18, "F").value = amount1
                .Cells(1, "T").value = TextBox6.Text
                .Cells(36, "F").value = amount                 'total amount

                .Cells(37, "F").value = total                  ' total taxes

                .Cells(39, "F").value = total + amount          'grand total

                .Cells(19, "B").value = ComboBox10.Text
            End If

            If ComboBox3.Text <> "" Then
                .Cells(21, "A").value = "2"
                .Cells(21, "B").value = ComboBox3.Text
                .Cells(21, "C").value = TextBox2.Text
                .Cells(21, "E").value = TextBox7.Text
                .Cells(21, "F").value = amount2
                .Cells(22, "B").value = ComboBox11.Text
                .Cells(2, "T").value = TextBox8.Text
            End If

            If ComboBox4.Text <> "" Then
                .Cells(24, "A").value = "3"
                .Cells(24, "B").value = ComboBox4.Text
                .Cells(24, "C").value = TextBox12.Text
                .Cells(24, "E").value = TextBox13.Text
                .Cells(24, "F").value = amount3
                .Cells(25, "B").value = ComboBox12.Text
                .Cells(3, "T").value = TextBox14.Text
            End If

            If ComboBox5.Text <> "" Then
                .Cells(27, "A").value = "4"
                .Cells(27, "B").value = ComboBox5.Text
                .Cells(27, "C").value = TextBox17.Text
                .Cells(27, "E").value = TextBox18.Text
                .Cells(27, "F").value = amount4
                .Cells(28, "B").value = ComboBox13.Text
                .Cells(4, "T").value = TextBox19.Text
            End If

            If ComboBox6.Text <> "" Then
                .Cells(30, "A").value = "5"
                .Cells(30, "B").value = ComboBox6.Text
                .Cells(30, "C").value = TextBox22.Text
                .Cells(30, "E").value = TextBox23.Text
                .Cells(30, "F").value = amount5
                .Cells(31, "B").value = ComboBox14.Text
                .Cells(5, "T").value = TextBox24.Text
            End If

            If ComboBox7.Text <> "" Then
                .Cells(33, "A").value = "6"
                .Cells(33, "B").value = ComboBox7.Text
                .Cells(33, "C").value = TextBox27.Text
                .Cells(33, "E").value = TextBox28.Text
                .Cells(33, "F").value = amount6
                .Cells(34, "B").value = ComboBox15.Text
                .Cells(6, "T").value = TextBox29.Text
            End If

            If ComboBox8.Text = "Delivery challan" Then
                .Cells(18, "E").value = ""
                .Cells(18, "F").value = ""
                .Cells(36, "F").value = ""                 'total amount

                .Cells(37, "F").value = ""                  ' total taxes

                .Cells(39, "F").value = ""        'grand total"
                .Cells(21, "E").value = ""
                .Cells(21, "F").value = ""
                .Cells(24, "E").value = ""
                .Cells(24, "F").value = ""
                .Cells(27, "E").value = ""
                .Cells(27, "F").value = ""
                .Cells(30, "E").value = ""
                .Cells(30, "F").value = ""
                .Cells(33, "E").value = ""
                .Cells(33, "F").value = ""

            End If
        End With

        '==============================================================================================

        With xlWorkSheet
            '~~> Directly type the values that we want
            Call movecell(0)
            .Range(con).Value = "                                                                          " + ComboBox8.SelectedItem + "                                               "
            movecell(1)
            .Range(con).Value = "              Aryan Enterprise        "
            movecell(1)
            .Range(con).Value = "                                                                             We serve home better"
            movecell(1)
            .Range(con).Value = "                                                                Address:- 107 road no-2 ,Sahar Villagae Andheri-E , Mumbai-99                                           "
            movecell(1)
            .Range(con).Value = "                                                                     Phone/Fax- 022-9221415392, Email- info@aryanchairs.com                                              "
            movecell(2)
            .Range(con).Value = "To,"
            movecell(1)
            .Range(con).Value = ComboBox1.Text
            movecell(1)
            '.Range(con).Value = TextBox9.Text
            movecell(-1)
            movecell(5)
            .Range(con).Value = "Dear Sir"
            movecell(1)
            .Range(con).Value = "We thank you for your keen interest in our products and services."
            movecell(2)
            .Range(con).Value = "SR.NO"
            movecell(3)
            ' .Range(con).Value = "1"
            If ComboBox3.Text = "" Then
                movecell(4)
                movecell(-4)
            Else
                movecell(4)
                '.Range(con).Value = "2"
                movecell(-4)
            End If

            movecolumn(1)
            movecell(-3)
            .Range(con).Value = "DESCRIPTION"
            movecell(3)
            ' .Range(con).Value = ComboBox2.Text
            If ComboBox3.Text = "" Then
                movecell(4)
                movecell(-4)
            Else
                movecell(4)
                '    .Range(con).Value = ComboBox3.Text
                movecell(-4)
            End If
            movecell(1)
            '.Range(con).Value = TextBox3.Text
            If ComboBox11.Text = "" Then
                movecell(4)
                movecell(-4)
            Else
                movecell(4)
                ' .Range(con).Value = TextBox1.Text
                movecell(-4)
            End If
            '.Cells(37, "B").value = "TAXES"
            movecell(17)
            .Range(con).Value = "VAT                                                                                                                                                                                 "

            movecolumn(1)
            movecell(-21)
            .Range(con).Value = "QTY."
            movecell(3)
            ' .Range(con).Value = TextBox4.Text
            If TextBox2.Text = "0" Then
                movecell(4)
                movecell(-4)
            Else
                movecell(4)
                '.Range(con).Value = TextBox2.Text
                movecell(-4)
            End If
            'movecell(-3)           row =19
            '.Range(con).Value = ""

            movecolumn(2)
            movecell(-12)
            '.Range(con).Value = b
            movecell(1)
            .Range(con).Value = "QTN : " + "/Neu/" + randomstr
            movecell(8)
            .Range(con).Value = "UNIT PRICE    "

            movecell(3)
            ' .Range(con).Value = TextBox5.Text
            If TextBox7.Text = "0" Then
                movecell(4)
                movecell(-4)
            Else
                movecell(4)
                '    .Range(con).Value = TextBox7.Text
                movecell(-4)
            End If
            movecolumn(-3)

            movecell(17)
            .Range(con).Value = "TOTAL"

            movecolumn(3)
            movecell(1) 'row=37 column E
            ' .Range(con).Value = TextBox6.Text + "%"
            If TextBox8.Text = "" Then
                movecell(4)
                movecell(-4)
            Else
                '  .Range(con).Value = TextBox8.Text + "%"
            End If

            movecolumn(1)
            movecell(-21)
            .Range(con).Value = "AMOUNT"
            movecell(3)
            '.Range(con).Value = tax1
            If TextBox7.Text = "" Or TextBox7.Text = "0" Then
                movecell(4)
                movecell(-4)
            Else
                movecell(4)
                ' .Range(con).Value = tax2
                movecell(-4)
            End If
            movecell(17)
            '.Range(con).Value = total
            movecell(1) 'row 37 'column F
            .Range(con).Value = tax1 + tax2
            movecell(2)
            '.Range(con).Value = total + tax
            movecolumn(-3)
            movecell(0) 'row 39 column C
            .Range(con).Value = "GRAND TOTAL"
            movecolumn(-2)
            movecell(2)
            .Range(con).Value = "TERMS & CONDITIONS:"
            movecell(1)  'row 42 column A
            .Range(con).Value = "1) Payment: Immediate"
            movecell(1)  'row 43 column A
            .Range(con).Value = "2) Octroi: As Applicable."
            movecolumn(2)
            movecell(0)  'row 43 column C
            .Range(con).Value = "                                 Thanking You"
            movecolumn(-2)
            movecell(1)  'row 44 column A
            .Range(con).Value = "3) Fright, Packing, Forwarding Charges: As Applicable."
            movecolumn(2)
            movecell(0)  'row 44 column C
            .Range(con).Value = "                              For, Aryan Enterprise"
            movecell(1)  'row 45 column C
            .Range(con).Value = "                              Authorised Signatory"
            movecolumn(-2)
            movecell(2)  'row 47 column A
            .Range(con).Value = "                                                                             We serve home better"

            '~~> Insert formulas
            '.Range("B6").Formula = "=Sum(B2:B5)"
            '.Range("B7").Formula = "=Average(B2:B5)"

            '~~> Shade the titles


            With .Range("A1:F1")
                ' .Interior.ColorIndex = 1 '<~~ Cell Back Color Black
                .VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
                '.HorizontalAlignment = Excel.XlVAlign.xlVAlignCenter
                .EntireRow.AutoFit()
                '.EntireRow.Justify()
                .ColumnWidth = 6
                With .Font()
                    .ColorIndex = 1 '<~~ Font Color White
                    .Size = 12
                    .Name = "Arial"
                    ' .Underline = Excel.XlUnderlineStyle.xlUnderlineStyleSingle
                    .Bold = True
                End With

                'Create Border 
                With .Borders(Excel.XlBordersIndex.xlEdgeLeft)
                    .LineStyle = Excel.XlLineStyle.xlDouble
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Excel.XlBorderWeight.xlThin
                End With
                With .Borders(Excel.XlBordersIndex.xlEdgeTop)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Excel.XlBorderWeight.xlThin
                End With
                With .Borders(Excel.XlBordersIndex.xlEdgeBottom)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Excel.XlBorderWeight.xlThin
                End With
                With .Borders(Excel.XlBordersIndex.xlEdgeRight)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Excel.XlBorderWeight.xlThin
                End With
                With .Borders(Excel.XlBordersIndex.xlInsideVertical)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Excel.XlBorderWeight.xlThin
                End With
                With .Borders(Excel.XlBordersIndex.xlInsideHorizontal)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Excel.XlBorderWeight.xlThin
                End With

            End With

            With .Range("F1:F47")
                With .Borders(Excel.XlBordersIndex.xlEdgeRight)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Excel.XlBorderWeight.xlMedium
                End With
            End With

            With .Range("A2")
                ' .Interior.ColorIndex = 1 '<~~ Cell Back Color Black
                .VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
                '.HorizontalAlignment = Excel.XlVAlign.xlVAlignCenter
                '.EntireRow.AutoFit()
                '.EntireRow.Justify()
                With .Font()
                    .ColorIndex = 1 '<~~ Font Color White
                    .Size = 36
                    .Name = "giro"
                    ' .Underline = Excel.XlUnderlineStyle.xlUnderlineStyleSingle
                    .Bold = True
                End With



            End With

            With .Range("A3:F3,A47:F47")
                .Interior.ColorIndex = 49 '<~~ Cell Back Color Black
                .VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
                '.HorizontalAlignment = Excel.XlVAlign.xlVAlignCenter
                .EntireRow.AutoFit()
                With .Font()
                    .ColorIndex = 2 '<~~ Font Color White
                    .Size = 10
                    .Name = "Arial"
                    ' .Underline = Excel.XlUnderlineStyle.xlUnderlineStyleSingle
                    '.Bold = True
                End With



            End With

            With .Range("A4:F4,A5:F5")
                ' .Interior.ColorIndex = 1 '<~~ Cell Back Color Black
                .VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
                '.HorizontalAlignment = Excel.XlVAlign.xlVAlignCenter
                .EntireRow.AutoFit()
                With .Font()
                    .ColorIndex = 49 '<~~ Font Color blue
                    .Size = 9
                    .Name = "Arial"
                    ' .Underline = Excel.XlUnderlineStyle.xlUnderlineStyleSingle
                    .Bold = True
                End With

                'Create Border 
                With .Borders(Excel.XlBordersIndex.xlEdgeLeft)
                    .LineStyle = Excel.XlLineStyle.xlDouble
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Excel.XlBorderWeight.xlThin
                End With
                With .Borders(Excel.XlBordersIndex.xlEdgeRight)
                    .LineStyle = Excel.XlLineStyle.xlDouble
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Excel.XlBorderWeight.xlThin
                End With
                With .Borders(Excel.XlBordersIndex.xlEdgeTop)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Excel.XlBorderWeight.xlThin
                End With
                With .Borders(Excel.XlBordersIndex.xlEdgeBottom)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Excel.XlBorderWeight.xlThin
                End With
                With .Borders(Excel.XlBordersIndex.xlInsideVertical)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Excel.XlBorderWeight.xlThin
                End With
                With .Borders(Excel.XlBordersIndex.xlInsideHorizontal)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Excel.XlBorderWeight.xlThin
                End With

            End With

            With .Range("A7:E7,A8:E8,A11:E11")
                ' .Interior.ColorIndex = 1 '<~~ Cell Back Color Black
                '.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
                ' .HorizontalAlignment = Excel.XlVAlign.xlVAlignCenter

                With .Font()
                    .ColorIndex = 1 '<~~ Font Color White
                    .Size = 10
                    .Name = "Arial"
                    ' .Underline = Excel.XlUnderlineStyle.xlUnderlineStyleSingle
                    .Bold = True
                End With



            End With


            With .Range("T1")
                .ColumnWidth = 0
            End With


            With .Range("B1")
                ' .Interior.ColorIndex = 1 '<~~ Cell Back Color Black
                .VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
                '.HorizontalAlignment = Excel.XlVAlign.xlVAlignCenter
                '.EntireRow.AutoFit()
                '.EntireRow.Justify()
                .ColumnWidth = 50

            End With

            With .Range("F1")
                ' .Interior.ColorIndex = 1 '<~~ Cell Back Color Black
                .VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
                '.HorizontalAlignment = Excel.XlVAlign.xlVAlignCenter
                '.EntireRow.AutoFit()
                '.EntireRow.Justify()
                .ColumnWidth = 20

            End With

            With .Range("A16:F16")
                .Interior.ColorIndex = 47 '<~~ Cell Back Color Black
                .VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
                .HorizontalAlignment = Excel.XlVAlign.xlVAlignCenter
                '.EntireRow.AutoFit()
                '.EntireRow.Justify()

                With .Font()
                    .ColorIndex = 2 '<~~ Font Color white
                    .Size = 9
                    .Name = "Arial"
                    ' .Underline = Excel.XlUnderlineStyle.xlUnderlineStyleSingle
                    .Bold = True
                End With

                'Create Border 
                With .Borders(Excel.XlBordersIndex.xlEdgeLeft)
                    .LineStyle = Excel.XlLineStyle.xlDouble
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Excel.XlBorderWeight.xlThin
                End With
                With .Borders(Excel.XlBordersIndex.xlEdgeTop)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Excel.XlBorderWeight.xlThin
                End With
                With .Borders(Excel.XlBordersIndex.xlEdgeBottom)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Excel.XlBorderWeight.xlThin
                End With
                With .Borders(Excel.XlBordersIndex.xlInsideVertical)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Excel.XlBorderWeight.xlThin
                End With
                With .Borders(Excel.XlBordersIndex.xlInsideHorizontal)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Excel.XlBorderWeight.xlThin
                End With


            End With

            With .Range("A17:A38")
                '.Interior.ColorIndex = 47 '<~~ Cell Back Color Black
                .VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
                .HorizontalAlignment = Excel.XlVAlign.xlVAlignCenter
                .EntireRow.AutoFit()
                '.EntireRow.Justify()

                With .Font()
                    .ColorIndex = 1 '<~~ Font Color white
                    .Size = 9
                    .Name = "Arial"
                    ' .Underline = Excel.XlUnderlineStyle.xlUnderlineStyleSingle
                    .Bold = True
                End With

                'Create Border 
                With .Borders(Excel.XlBordersIndex.xlEdgeRight)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Excel.XlBorderWeight.xlThin
                End With

            End With

            With .Range("B17:B35,C17:C35,F17:F37")
                '.Interior.ColorIndex = 47 '<~~ Cell Back Color Black
                '.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
                '.HorizontalAlignment = Excel.XlVAlign.xlVAlignCenter
                .EntireRow.AutoFit()
                '.EntireRow.Justify()

                With .Font()
                    .ColorIndex = 1 '<~~ Font Color white
                    .Size = 10
                    .Name = "Arial"
                    ' .Underline = Excel.XlUnderlineStyle.xlUnderlineStyleSingle
                    '.Bold = True
                End With

                'Create Border 
                With .Borders(Excel.XlBordersIndex.xlEdgeRight)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Excel.XlBorderWeight.xlThin
                End With
                With .Borders(Excel.XlBordersIndex.xlEdgeLeft)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Excel.XlBorderWeight.xlThin
                End With
            End With

            With .Cells(37, "F")
                With .Font()
                    .ColorIndex = 0 '<~~ Font Color white
                    .Size = 10
                    .Name = "Arial"
                    ' .Underline = Excel.XlUnderlineStyle.xlUnderlineStyleSingle
                    '.Bold = True
                End With

                'Create Border 

                With .Borders(Excel.XlBordersIndex.xlEdgeLeft)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Excel.XlBorderWeight.xlThin
                End With
                With .Borders(Excel.XlBordersIndex.xlInsideVertical)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Excel.XlBorderWeight.xlThin
                End With
                With .Borders(Excel.XlBordersIndex.xlInsideHorizontal)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Excel.XlBorderWeight.xlThin
                End With
            End With

            With .Range("B36:B38,C36:C38,D36:D38,E36:E38,F36:F38")
                '.Interior.ColorIndex = 47 '<~~ Cell Back Color Black
                '.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
                '.HorizontalAlignment = Excel.XlVAlign.xlVAlignCenter
                .EntireRow.AutoFit()
                '.EntireRow.Justify()

                With .Font()
                    .ColorIndex = 1 '<~~ Font Color white
                    .Size = 10
                    .Name = "Arial"
                    ' .Underline = Excel.XlUnderlineStyle.xlUnderlineStyleSingle
                    '.Bold = True
                End With

                'Create Border 
                With .Borders(Excel.XlBordersIndex.xlEdgeTop)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Excel.XlBorderWeight.xlThin
                End With

                With .Borders(Excel.XlBordersIndex.xlEdgeLeft)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Excel.XlBorderWeight.xlThin
                End With
                With .Borders(Excel.XlBordersIndex.xlInsideVertical)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Excel.XlBorderWeight.xlThin
                End With
                With .Borders(Excel.XlBordersIndex.xlInsideHorizontal)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Excel.XlBorderWeight.xlThin
                End With
            End With

            With .Range("A39:F39,A40:F40")
                .Interior.ColorIndex = 47 '<~~ Cell Back Color Black
                '.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
                '.HorizontalAlignment = Excel.XlVAlign.xlVAlignCenter
                '.EntireRow.AutoFit()
                '.EntireRow.Justify()

                With .Font()
                    .ColorIndex = 2 '<~~ Font Color white
                    .Size = 9
                    .Name = "Arial"
                    ' .Underline = Excel.XlUnderlineStyle.xlUnderlineStyleSingle
                    .Bold = True
                End With

            End With

            With .Range("A41:A44,C43:C45")
                '.Interior.ColorIndex = 47 '<~~ Cell Back Color Black
                '.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
                '.HorizontalAlignment = Excel.XlVAlign.xlVAlignCenter
                .EntireRow.AutoFit()
                '.EntireRow.Justify()

                With .Font()
                    .ColorIndex = 1 '<~~ Font Color white
                    .Size = 8
                    .Name = "Arial"
                    ' .Underline = Excel.XlUnderlineStyle.xlUnderlineStyleSingle
                    .Bold = True
                End With

            End With


            With .Range("B41:B45")
                With .Borders(Excel.XlBordersIndex.xlEdgeRight)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Excel.XlBorderWeight.xlThin
                End With
            End With


            With .Range("A46:F46")
                With .Borders(Excel.XlBordersIndex.xlEdgeTop)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Excel.XlBorderWeight.xlMedium
                End With
            End With



            ' ~~> Set the number format
            '.Range("E7").NumberFormat = "##/##/#####"
            ' .Range ("E7").DisplayFormat =""

            '~~> Autofitting text in columns
            '.Columns("A").Entire Column.AutoFit()
            '.Rows("2").EntireRow.Autofit()
        End With

    End Sub
    Private Sub savework()
        Dim sec As String
        Dim dat As String
        Dim hh As String
        Dim mm As String
        Dim alternate As String
        dat = Date.Today.ToString("dd_MM_yy")
        hh = System.DateTime.Now.Hour
        mm = System.DateTime.Now.Minute
        sec = System.DateTime.Now.Second
        alternate = ""
        If (ComboBox1.Text = "") Then
            alternate = "user_" + dat + "_" + hh + "_" + mm + "_" + sec
        Else
            alternate = ComboBox1.Text + dat + "_" + hh + "_" + mm + "_" + sec
        End If
        filename = "C:\shree\" + alternate

        xlWorkBook.SaveAs(Filename:=filename)
    End Sub

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load


        'Default values
        ComboBox8.Text = "Quotation"

        'file operation
        'create file

        If File.Exists(path) = False Then
            ' Create a file to write to. 
            Dim createText() As String = {""}
            File.WriteAllLines(path, createText)
        End If

        If File.Exists(pathitems) = False Then
            ' Create a file to write to. 
            Dim createText() As String = {""}
            File.WriteAllLines(pathitems, createText)
        End If

        'load values from textfile to combobox
        ComboBox1.Items.AddRange(System.IO.File.ReadAllLines(path))
        ComboBox2.Items.AddRange(System.IO.File.ReadAllLines(pathitems))
        ComboBox3.Items.AddRange(System.IO.File.ReadAllLines(pathitems))
        ComboBox4.Items.AddRange(System.IO.File.ReadAllLines(pathitems))
        ComboBox5.Items.AddRange(System.IO.File.ReadAllLines(pathitems))
        ComboBox6.Items.AddRange(System.IO.File.ReadAllLines(pathitems))
        ComboBox7.Items.AddRange(System.IO.File.ReadAllLines(pathitems))

        ComboBox9.Items.AddRange(System.IO.File.ReadAllLines(place))

        ComboBox10.Items.AddRange(System.IO.File.ReadAllLines(desc))
        ComboBox11.Items.AddRange(System.IO.File.ReadAllLines(desc))
        ComboBox12.Items.AddRange(System.IO.File.ReadAllLines(desc))
        ComboBox13.Items.AddRange(System.IO.File.ReadAllLines(desc))
        ComboBox14.Items.AddRange(System.IO.File.ReadAllLines(desc))
        ComboBox15.Items.AddRange(System.IO.File.ReadAllLines(desc))

    End Sub


    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Button1.Enabled = True
        Button2.Enabled = False
        Button3.Enabled = False

        Dim con As New OleDb.OleDbConnection
        Dim dbProvider As String
        Dim dbSource As String
        Dim ds As New DataSet

        Dim sql As String


        dbProvider = "Provider=Microsoft.ACE.OLEDB.12.0;"
        dbSource = "Data Source = C:/shree/files/user.accdb"

        con.ConnectionString = dbProvider & dbSource
        con.Open()
        sql = "INSERT INTO users (cust_name,[number],gen_date,category) values ('" & ComboBox1.Text & "','" & "/Neu/" + randomstr & "','" & DateTimePicker1.Value.ToString("MMM dd, yyyy") & "','" & ComboBox8.Text & "')"

        Dim run = New OleDb.OleDbCommand

        run = New OleDbCommand(sql, con)

        run.ExecuteNonQuery()

        'MsgBox("New Record added to the Database")
        con.Close()

        savework()
        'For Each w In xlApp.Workbooks
        '    w.Save()
        'Next w
        xlApp.Workbooks.Close()
        xlApp.Quit()

    End Sub


    Private Sub FlowLayoutPanel1_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs)

    End Sub

    Private Sub TextBox6_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox6.TextChanged


    End Sub

    Private Sub TextBox8_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox8.TextChanged

    End Sub

    Private Sub TextBox11_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub EditToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles EditToolStripMenuItem.Click
        Form3.Show()
        'Me.Close()
    End Sub

    Private Sub ExitToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExitToolStripMenuItem.Click
        Me.Close()
    End Sub

    Private Sub AllToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AllToolStripMenuItem.Click
        Form4.Show()
        'Me.Close()
    End Sub

    Private Sub Button3_Click(ByVal sender As Object, ByVal e As System.EventArgs)

        'xlWorkBookuser = xlAppuser.Workbooks.Open("c:\shree\user.xlsx")
        'xlWorkSheetuser = CType(xlWorkBookuser.Worksheets.Item("Sheet1"), Excel.Worksheet)
        'With xlWorkSheetuser
        '    .Cells(11, "E").value = "ll"      'date

        'End With
        'With xlWorkSheetuser
        '    'Call movecell(0)
        '    .Range("A1").Value = "dell"
        '    ' Call movecolumn(1)
        '    ' .Range(con).Value = "/Neu/" '+ randomstr
        'End With
        'xlWorkBookuser.Save()
        'xlWorkBookuser.Close()
        'xlAppuser.Quit()
    End Sub

    Private Sub TextBox9_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub Button3_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Button1.Enabled = True
        Button2.Enabled = False
        Button3.Enabled = False
        xlWorkBook.Close(SaveChanges:=False)
        xlApp.Quit()

    End Sub
End Class
