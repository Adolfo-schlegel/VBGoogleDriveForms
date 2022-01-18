<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Form1
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
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
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Form1))
        Me.btnBuscar = New System.Windows.Forms.Button()
        Me.btnSubir = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.FilePath = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'btnBuscar
        '
        Me.btnBuscar.Location = New System.Drawing.Point(12, 12)
        Me.btnBuscar.Name = "btnBuscar"
        Me.btnBuscar.Size = New System.Drawing.Size(133, 23)
        Me.btnBuscar.TabIndex = 0
        Me.btnBuscar.Text = "Buscar Archivo"
        Me.btnBuscar.UseVisualStyleBackColor = True
        '
        'btnSubir
        '
        Me.btnSubir.Location = New System.Drawing.Point(398, 218)
        Me.btnSubir.Name = "btnSubir"
        Me.btnSubir.Size = New System.Drawing.Size(115, 23)
        Me.btnSubir.TabIndex = 1
        Me.btnSubir.Text = "Subir"
        Me.btnSubir.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(20, 38)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(119, 15)
        Me.Label1.TabIndex = 2
        Me.Label1.Text = "Ruta de tu archivo ->"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(145, 38)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(0, 15)
        Me.Label2.TabIndex = 3
        '
        'FilePath
        '
        Me.FilePath.AutoSize = True
        Me.FilePath.Location = New System.Drawing.Point(145, 38)
        Me.FilePath.Name = "FilePath"
        Me.FilePath.Size = New System.Drawing.Size(0, 15)
        Me.FilePath.TabIndex = 4
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 15.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.ActiveCaption
        Me.BackgroundImage = CType(resources.GetObject("$this.BackgroundImage"), System.Drawing.Image)
        Me.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.ClientSize = New System.Drawing.Size(525, 253)
        Me.Controls.Add(Me.FilePath)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.btnSubir)
        Me.Controls.Add(Me.btnBuscar)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "Form1"
        Me.Text = "Google Drive File"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents btnBuscar As Button
    Friend WithEvents btnSubir As Button
    Friend WithEvents Label1 As Label
    Friend WithEvents Label2 As Label
    Friend WithEvents FilePath As Label
End Class
