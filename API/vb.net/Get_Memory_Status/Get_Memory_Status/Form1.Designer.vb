<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form1
    Inherits System.Windows.Forms.Form

    'Form 覆寫 Dispose 以清除元件清單。
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

    '為 Windows Form 設計工具的必要項
    Private components As System.ComponentModel.IContainer

    '注意: 以下為 Windows Form 設計工具所需的程序
    '可以使用 Windows Form 設計工具進行修改。
    '請勿使用程式碼編輯器進行修改。
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.TextBox1 = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.TextBox2 = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.TextBox3 = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.TextBox4 = New System.Windows.Forms.TextBox()
        Me.TextBox5 = New System.Windows.Forms.TextBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.TextBox6 = New System.Windows.Forms.TextBox()
        Me.TextBox7 = New System.Windows.Forms.TextBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.TextBox8 = New System.Windows.Forms.TextBox()
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.SuspendLayout()
        '
        'TextBox1
        '
        Me.TextBox1.Font = New System.Drawing.Font("新細明體", 16.0!)
        Me.TextBox1.Location = New System.Drawing.Point(182, 28)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(151, 33)
        Me.TextBox1.TabIndex = 0
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("新細明體", 16.0!)
        Me.Label1.Location = New System.Drawing.Point(34, 31)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(142, 22)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "記憶體使用率"
        '
        'TextBox2
        '
        Me.TextBox2.Font = New System.Drawing.Font("新細明體", 16.0!)
        Me.TextBox2.Location = New System.Drawing.Point(182, 86)
        Me.TextBox2.Name = "TextBox2"
        Me.TextBox2.Size = New System.Drawing.Size(151, 33)
        Me.TextBox2.TabIndex = 2
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("新細明體", 16.0!)
        Me.Label2.Location = New System.Drawing.Point(12, 89)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(164, 22)
        Me.Label2.TabIndex = 3
        Me.Label2.Text = "物理記憶體大小"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("新細明體", 16.0!)
        Me.Label3.Location = New System.Drawing.Point(12, 143)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(164, 22)
        Me.Label3.TabIndex = 3
        Me.Label3.Text = "可用記憶體大小"
        '
        'TextBox3
        '
        Me.TextBox3.Font = New System.Drawing.Font("新細明體", 16.0!)
        Me.TextBox3.Location = New System.Drawing.Point(182, 140)
        Me.TextBox3.Name = "TextBox3"
        Me.TextBox3.Size = New System.Drawing.Size(151, 33)
        Me.TextBox3.TabIndex = 2
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Font = New System.Drawing.Font("新細明體", 16.0!)
        Me.Label4.Location = New System.Drawing.Point(397, 31)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(142, 22)
        Me.Label4.TabIndex = 3
        Me.Label4.Text = "總物理加虛擬"
        '
        'TextBox4
        '
        Me.TextBox4.Font = New System.Drawing.Font("新細明體", 16.0!)
        Me.TextBox4.Location = New System.Drawing.Point(545, 28)
        Me.TextBox4.Name = "TextBox4"
        Me.TextBox4.Size = New System.Drawing.Size(151, 33)
        Me.TextBox4.TabIndex = 2
        '
        'TextBox5
        '
        Me.TextBox5.Font = New System.Drawing.Font("新細明體", 16.0!)
        Me.TextBox5.Location = New System.Drawing.Point(545, 84)
        Me.TextBox5.Name = "TextBox5"
        Me.TextBox5.Size = New System.Drawing.Size(151, 33)
        Me.TextBox5.TabIndex = 2
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Font = New System.Drawing.Font("新細明體", 16.0!)
        Me.Label5.Location = New System.Drawing.Point(419, 89)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(120, 22)
        Me.Label5.TabIndex = 3
        Me.Label5.Text = "可用分頁檔"
        '
        'TextBox6
        '
        Me.TextBox6.Font = New System.Drawing.Font("新細明體", 16.0!)
        Me.TextBox6.Location = New System.Drawing.Point(182, 197)
        Me.TextBox6.Name = "TextBox6"
        Me.TextBox6.Size = New System.Drawing.Size(151, 33)
        Me.TextBox6.TabIndex = 2
        '
        'TextBox7
        '
        Me.TextBox7.Font = New System.Drawing.Font("新細明體", 16.0!)
        Me.TextBox7.Location = New System.Drawing.Point(182, 251)
        Me.TextBox7.Name = "TextBox7"
        Me.TextBox7.Size = New System.Drawing.Size(151, 33)
        Me.TextBox7.TabIndex = 2
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Font = New System.Drawing.Font("新細明體", 16.0!)
        Me.Label6.Location = New System.Drawing.Point(56, 200)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(120, 22)
        Me.Label6.TabIndex = 3
        Me.Label6.Text = "總虛擬大小"
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Font = New System.Drawing.Font("新細明體", 16.0!)
        Me.Label7.Location = New System.Drawing.Point(34, 254)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(142, 22)
        Me.Label7.TabIndex = 3
        Me.Label7.Text = "可用虛擬大小"
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Font = New System.Drawing.Font("新細明體", 16.0!)
        Me.Label8.Location = New System.Drawing.Point(56, 306)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(120, 22)
        Me.Label8.TabIndex = 3
        Me.Label8.Text = "可擴展虛擬"
        '
        'TextBox8
        '
        Me.TextBox8.Font = New System.Drawing.Font("新細明體", 16.0!)
        Me.TextBox8.Location = New System.Drawing.Point(182, 303)
        Me.TextBox8.Name = "TextBox8"
        Me.TextBox8.Size = New System.Drawing.Size(151, 33)
        Me.TextBox8.TabIndex = 2
        '
        'Timer1
        '
        Me.Timer1.Enabled = True
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(800, 450)
        Me.Controls.Add(Me.Label8)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.TextBox8)
        Me.Controls.Add(Me.TextBox5)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.TextBox4)
        Me.Controls.Add(Me.TextBox7)
        Me.Controls.Add(Me.TextBox3)
        Me.Controls.Add(Me.TextBox6)
        Me.Controls.Add(Me.TextBox2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.TextBox1)
        Me.Name = "Form1"
        Me.Text = "Form1"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents TextBox1 As TextBox
    Friend WithEvents Label1 As Label
    Friend WithEvents TextBox2 As TextBox
    Friend WithEvents Label2 As Label
    Friend WithEvents Label3 As Label
    Friend WithEvents TextBox3 As TextBox
    Friend WithEvents Label4 As Label
    Friend WithEvents TextBox4 As TextBox
    Friend WithEvents TextBox5 As TextBox
    Friend WithEvents Label5 As Label
    Friend WithEvents TextBox6 As TextBox
    Friend WithEvents TextBox7 As TextBox
    Friend WithEvents Label6 As Label
    Friend WithEvents Label7 As Label
    Friend WithEvents Label8 As Label
    Friend WithEvents TextBox8 As TextBox
    Friend WithEvents Timer1 As Timer
End Class
