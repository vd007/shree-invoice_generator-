<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form4
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim Cust_nameLabel As System.Windows.Forms.Label
        Dim NumberLabel As System.Windows.Forms.Label
        Me.Cust_nameTextBox = New System.Windows.Forms.TextBox()
        Me.NumberTextBox = New System.Windows.Forms.TextBox()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.TextBox1 = New System.Windows.Forms.TextBox()
        Me.TextBox2 = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.TextBox3 = New System.Windows.Forms.TextBox()
        Me.TextBox4 = New System.Windows.Forms.TextBox()
        Me.TextBox5 = New System.Windows.Forms.TextBox()
        Me.TextBox6 = New System.Windows.Forms.TextBox()
        Me.TextBox7 = New System.Windows.Forms.TextBox()
        Me.TextBox8 = New System.Windows.Forms.TextBox()
        Me.TextBox9 = New System.Windows.Forms.TextBox()
        Me.TextBox10 = New System.Windows.Forms.TextBox()
        Me.TextBox11 = New System.Windows.Forms.TextBox()
        Me.TextBox12 = New System.Windows.Forms.TextBox()
        Me.TextBox13 = New System.Windows.Forms.TextBox()
        Me.TextBox14 = New System.Windows.Forms.TextBox()
        Me.TextBox15 = New System.Windows.Forms.TextBox()
        Me.TextBox16 = New System.Windows.Forms.TextBox()
        Me.TextBox17 = New System.Windows.Forms.TextBox()
        Me.TextBox18 = New System.Windows.Forms.TextBox()
        Cust_nameLabel = New System.Windows.Forms.Label()
        NumberLabel = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'Cust_nameLabel
        '
        Cust_nameLabel.AutoSize = True
        Cust_nameLabel.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Cust_nameLabel.Location = New System.Drawing.Point(74, 41)
        Cust_nameLabel.Name = "Cust_nameLabel"
        Cust_nameLabel.Size = New System.Drawing.Size(105, 16)
        Cust_nameLabel.TabIndex = 1
        Cust_nameLabel.Text = "Customer Name"
        AddHandler Cust_nameLabel.Click, AddressOf Me.Cust_nameLabel_Click
        '
        'NumberLabel
        '
        NumberLabel.AutoSize = True
        NumberLabel.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        NumberLabel.Location = New System.Drawing.Point(311, 41)
        NumberLabel.Name = "NumberLabel"
        NumberLabel.Size = New System.Drawing.Size(56, 16)
        NumberLabel.TabIndex = 3
        NumberLabel.Text = "Number"
        AddHandler NumberLabel.Click, AddressOf Me.NumberLabel_Click
        '
        'Cust_nameTextBox
        '
        Me.Cust_nameTextBox.Location = New System.Drawing.Point(77, 77)
        Me.Cust_nameTextBox.Name = "Cust_nameTextBox"
        Me.Cust_nameTextBox.Size = New System.Drawing.Size(100, 20)
        Me.Cust_nameTextBox.TabIndex = 2
        '
        'NumberTextBox
        '
        Me.NumberTextBox.Location = New System.Drawing.Point(277, 77)
        Me.NumberTextBox.Name = "NumberTextBox"
        Me.NumberTextBox.Size = New System.Drawing.Size(100, 20)
        Me.NumberTextBox.TabIndex = 4
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(439, 299)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(74, 33)
        Me.Button1.TabIndex = 9
        Me.Button1.Text = "next"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'Button2
        '
        Me.Button2.Location = New System.Drawing.Point(326, 298)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(71, 34)
        Me.Button2.TabIndex = 10
        Me.Button2.Text = "previous"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(508, 41)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(37, 16)
        Me.Label1.TabIndex = 11
        Me.Label1.Text = "Date"
        '
        'TextBox1
        '
        Me.TextBox1.Location = New System.Drawing.Point(492, 76)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(100, 20)
        Me.TextBox1.TabIndex = 12
        '
        'TextBox2
        '
        Me.TextBox2.Location = New System.Drawing.Point(693, 76)
        Me.TextBox2.Name = "TextBox2"
        Me.TextBox2.Size = New System.Drawing.Size(100, 20)
        Me.TextBox2.TabIndex = 13
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(719, 41)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(40, 16)
        Me.Label2.TabIndex = 14
        Me.Label2.Text = "Type"
        '
        'TextBox3
        '
        Me.TextBox3.Location = New System.Drawing.Point(77, 112)
        Me.TextBox3.Name = "TextBox3"
        Me.TextBox3.Size = New System.Drawing.Size(100, 20)
        Me.TextBox3.TabIndex = 15
        '
        'TextBox4
        '
        Me.TextBox4.Location = New System.Drawing.Point(277, 112)
        Me.TextBox4.Name = "TextBox4"
        Me.TextBox4.Size = New System.Drawing.Size(100, 20)
        Me.TextBox4.TabIndex = 16
        '
        'TextBox5
        '
        Me.TextBox5.Location = New System.Drawing.Point(494, 110)
        Me.TextBox5.Name = "TextBox5"
        Me.TextBox5.Size = New System.Drawing.Size(98, 20)
        Me.TextBox5.TabIndex = 17
        '
        'TextBox6
        '
        Me.TextBox6.Location = New System.Drawing.Point(693, 110)
        Me.TextBox6.Name = "TextBox6"
        Me.TextBox6.Size = New System.Drawing.Size(100, 20)
        Me.TextBox6.TabIndex = 18
        '
        'TextBox7
        '
        Me.TextBox7.Location = New System.Drawing.Point(77, 156)
        Me.TextBox7.Name = "TextBox7"
        Me.TextBox7.Size = New System.Drawing.Size(100, 20)
        Me.TextBox7.TabIndex = 19
        '
        'TextBox8
        '
        Me.TextBox8.Location = New System.Drawing.Point(277, 156)
        Me.TextBox8.Name = "TextBox8"
        Me.TextBox8.Size = New System.Drawing.Size(100, 20)
        Me.TextBox8.TabIndex = 20
        '
        'TextBox9
        '
        Me.TextBox9.Location = New System.Drawing.Point(494, 156)
        Me.TextBox9.Name = "TextBox9"
        Me.TextBox9.Size = New System.Drawing.Size(99, 20)
        Me.TextBox9.TabIndex = 21
        '
        'TextBox10
        '
        Me.TextBox10.Location = New System.Drawing.Point(693, 156)
        Me.TextBox10.Name = "TextBox10"
        Me.TextBox10.Size = New System.Drawing.Size(99, 20)
        Me.TextBox10.TabIndex = 22
        '
        'TextBox11
        '
        Me.TextBox11.Location = New System.Drawing.Point(77, 195)
        Me.TextBox11.Name = "TextBox11"
        Me.TextBox11.Size = New System.Drawing.Size(100, 20)
        Me.TextBox11.TabIndex = 23
        '
        'TextBox12
        '
        Me.TextBox12.Location = New System.Drawing.Point(277, 195)
        Me.TextBox12.Name = "TextBox12"
        Me.TextBox12.Size = New System.Drawing.Size(99, 20)
        Me.TextBox12.TabIndex = 24
        '
        'TextBox13
        '
        Me.TextBox13.Location = New System.Drawing.Point(492, 195)
        Me.TextBox13.Name = "TextBox13"
        Me.TextBox13.Size = New System.Drawing.Size(99, 20)
        Me.TextBox13.TabIndex = 25
        '
        'TextBox14
        '
        Me.TextBox14.Location = New System.Drawing.Point(693, 195)
        Me.TextBox14.Name = "TextBox14"
        Me.TextBox14.Size = New System.Drawing.Size(100, 20)
        Me.TextBox14.TabIndex = 26
        '
        'TextBox15
        '
        Me.TextBox15.Location = New System.Drawing.Point(77, 237)
        Me.TextBox15.Name = "TextBox15"
        Me.TextBox15.Size = New System.Drawing.Size(100, 20)
        Me.TextBox15.TabIndex = 27
        '
        'TextBox16
        '
        Me.TextBox16.Location = New System.Drawing.Point(277, 237)
        Me.TextBox16.Name = "TextBox16"
        Me.TextBox16.Size = New System.Drawing.Size(97, 20)
        Me.TextBox16.TabIndex = 28
        '
        'TextBox17
        '
        Me.TextBox17.Location = New System.Drawing.Point(496, 238)
        Me.TextBox17.Name = "TextBox17"
        Me.TextBox17.Size = New System.Drawing.Size(94, 20)
        Me.TextBox17.TabIndex = 29
        '
        'TextBox18
        '
        Me.TextBox18.Location = New System.Drawing.Point(695, 237)
        Me.TextBox18.Name = "TextBox18"
        Me.TextBox18.Size = New System.Drawing.Size(97, 20)
        Me.TextBox18.TabIndex = 30
        '
        'Form4
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(884, 562)
        Me.Controls.Add(Me.TextBox18)
        Me.Controls.Add(Me.TextBox17)
        Me.Controls.Add(Me.TextBox16)
        Me.Controls.Add(Me.TextBox15)
        Me.Controls.Add(Me.TextBox14)
        Me.Controls.Add(Me.TextBox13)
        Me.Controls.Add(Me.TextBox12)
        Me.Controls.Add(Me.TextBox11)
        Me.Controls.Add(Me.TextBox10)
        Me.Controls.Add(Me.TextBox9)
        Me.Controls.Add(Me.TextBox8)
        Me.Controls.Add(Me.TextBox7)
        Me.Controls.Add(Me.TextBox6)
        Me.Controls.Add(Me.TextBox5)
        Me.Controls.Add(Me.TextBox4)
        Me.Controls.Add(Me.TextBox3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.TextBox2)
        Me.Controls.Add(Me.TextBox1)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(NumberLabel)
        Me.Controls.Add(Me.NumberTextBox)
        Me.Controls.Add(Cust_nameLabel)
        Me.Controls.Add(Me.Cust_nameTextBox)
        Me.Name = "Form4"
        Me.Text = "Neurals ::Shree"
        Me.TopMost = True
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Cust_nameTextBox As System.Windows.Forms.TextBox
    Friend WithEvents NumberTextBox As System.Windows.Forms.TextBox
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox2 As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents TextBox3 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox4 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox5 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox6 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox7 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox8 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox9 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox10 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox11 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox12 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox13 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox14 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox15 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox16 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox17 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox18 As System.Windows.Forms.TextBox
End Class
