<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form1
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
        Me.btn00 = New System.Windows.Forms.Button()
        Me.btn01 = New System.Windows.Forms.Button()
        Me.btn02 = New System.Windows.Forms.Button()
        Me.btn12 = New System.Windows.Forms.Button()
        Me.btn11 = New System.Windows.Forms.Button()
        Me.btn10 = New System.Windows.Forms.Button()
        Me.btn20 = New System.Windows.Forms.Button()
        Me.btn21 = New System.Windows.Forms.Button()
        Me.btn22 = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'btn00
        '
        Me.btn00.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btn00.Location = New System.Drawing.Point(12, 13)
        Me.btn00.Name = "btn00"
        Me.btn00.Size = New System.Drawing.Size(57, 52)
        Me.btn00.TabIndex = 0
        Me.btn00.Tag = "0,0"
        Me.btn00.UseVisualStyleBackColor = True
        '
        'btn01
        '
        Me.btn01.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btn01.Location = New System.Drawing.Point(76, 13)
        Me.btn01.Name = "btn01"
        Me.btn01.Size = New System.Drawing.Size(57, 52)
        Me.btn01.TabIndex = 0
        Me.btn01.Tag = "0,1"
        Me.btn01.UseVisualStyleBackColor = True
        '
        'btn02
        '
        Me.btn02.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btn02.Location = New System.Drawing.Point(140, 13)
        Me.btn02.Name = "btn02"
        Me.btn02.Size = New System.Drawing.Size(57, 52)
        Me.btn02.TabIndex = 0
        Me.btn02.Tag = "0,2"
        Me.btn02.UseVisualStyleBackColor = True
        '
        'btn12
        '
        Me.btn12.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btn12.Location = New System.Drawing.Point(140, 69)
        Me.btn12.Name = "btn12"
        Me.btn12.Size = New System.Drawing.Size(57, 52)
        Me.btn12.TabIndex = 0
        Me.btn12.Tag = "1,2"
        Me.btn12.UseVisualStyleBackColor = True
        '
        'btn11
        '
        Me.btn11.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btn11.Location = New System.Drawing.Point(76, 69)
        Me.btn11.Name = "btn11"
        Me.btn11.Size = New System.Drawing.Size(57, 52)
        Me.btn11.TabIndex = 0
        Me.btn11.Tag = "1,1"
        Me.btn11.UseVisualStyleBackColor = True
        '
        'btn10
        '
        Me.btn10.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btn10.Location = New System.Drawing.Point(12, 69)
        Me.btn10.Name = "btn10"
        Me.btn10.Size = New System.Drawing.Size(57, 52)
        Me.btn10.TabIndex = 0
        Me.btn10.Tag = "1,0"
        Me.btn10.UseVisualStyleBackColor = True
        '
        'btn20
        '
        Me.btn20.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btn20.Location = New System.Drawing.Point(12, 125)
        Me.btn20.Name = "btn20"
        Me.btn20.Size = New System.Drawing.Size(57, 52)
        Me.btn20.TabIndex = 0
        Me.btn20.Tag = "2,0"
        Me.btn20.UseVisualStyleBackColor = True
        '
        'btn21
        '
        Me.btn21.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btn21.Location = New System.Drawing.Point(76, 125)
        Me.btn21.Name = "btn21"
        Me.btn21.Size = New System.Drawing.Size(57, 52)
        Me.btn21.TabIndex = 0
        Me.btn21.Tag = "2,1"
        Me.btn21.UseVisualStyleBackColor = True
        '
        'btn22
        '
        Me.btn22.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btn22.Location = New System.Drawing.Point(140, 125)
        Me.btn22.Name = "btn22"
        Me.btn22.Size = New System.Drawing.Size(57, 52)
        Me.btn22.TabIndex = 0
        Me.btn22.Tag = "2,2"
        Me.btn22.UseVisualStyleBackColor = True
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(219, 199)
        Me.Controls.Add(Me.btn22)
        Me.Controls.Add(Me.btn21)
        Me.Controls.Add(Me.btn20)
        Me.Controls.Add(Me.btn10)
        Me.Controls.Add(Me.btn11)
        Me.Controls.Add(Me.btn12)
        Me.Controls.Add(Me.btn02)
        Me.Controls.Add(Me.btn01)
        Me.Controls.Add(Me.btn00)
        Me.Name = "Form1"
        Me.Text = "Tic-Tac-Toe"
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents btn00 As System.Windows.Forms.Button
    Friend WithEvents btn01 As System.Windows.Forms.Button
    Friend WithEvents btn02 As System.Windows.Forms.Button
    Friend WithEvents btn12 As System.Windows.Forms.Button
    Friend WithEvents btn11 As System.Windows.Forms.Button
    Friend WithEvents btn10 As System.Windows.Forms.Button
    Friend WithEvents btn20 As System.Windows.Forms.Button
    Friend WithEvents btn21 As System.Windows.Forms.Button
    Friend WithEvents btn22 As System.Windows.Forms.Button

End Class
