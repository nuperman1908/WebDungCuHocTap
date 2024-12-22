using System.ComponentModel;

namespace TestingProgram;

partial class FormAddProduct_Test
{
    /// <summary>
    /// Required designer variable.
    /// </summary>
    private IContainer components = null;

    /// <summary>
    /// Clean up any resources being used.
    /// </summary>
    /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
    protected override void Dispose(bool disposing)
    {
        if (disposing && (components != null))
        {
            components.Dispose();
        }

        base.Dispose(disposing);
    }

    #region Windows Form Designer generated code

    /// <summary>
    /// Required method for Designer support - do not modify
    /// the contents of this method with the code editor.
    /// </summary>
    private void InitializeComponent()
    {
        lblTitle = new System.Windows.Forms.Label();
        lblTestCase = new System.Windows.Forms.Label();
        btnRun = new System.Windows.Forms.Button();
        txtTestCase = new System.Windows.Forms.TextBox();
        SuspendLayout();
        // 
        // lblTitle
        // 
        lblTitle.AutoSize = true;
        lblTitle.Font = new System.Drawing.Font("Arial", 14F, System.Drawing.FontStyle.Bold);
        lblTitle.Location = new System.Drawing.Point(150, 20);
        lblTitle.Name = "lblTitle";
        lblTitle.Size = new System.Drawing.Size(302, 22);
        lblTitle.TabIndex = 0;
        lblTitle.Text = "Test chức năng thêm sản phẩm";
        // 
        // lblTestCase
        // 
        lblTestCase.AutoSize = true;
        lblTestCase.Location = new System.Drawing.Point(50, 72);
        lblTestCase.Name = "lblTestCase";
        lblTestCase.Size = new System.Drawing.Size(133, 15);
        lblTestCase.TabIndex = 1;
        lblTestCase.Text = "Nhập số lượng test case";
        // 
        // btnRun
        // 
        btnRun.Location = new System.Drawing.Point(50, 128);
        btnRun.Name = "btnRun";
        btnRun.Size = new System.Drawing.Size(75, 23);
        btnRun.TabIndex = 5;
        btnRun.Text = "Chạy";
        btnRun.UseVisualStyleBackColor = true;
        btnRun.Click += btnRun_Click;
        // 
        // txtTestCase
        // 
        txtTestCase.Location = new System.Drawing.Point(50, 90);
        txtTestCase.Name = "txtTestCase";
        txtTestCase.Size = new System.Drawing.Size(200, 23);
        txtTestCase.TabIndex = 2;
        txtTestCase.TextChanged += TxtTestCase_TextChanged;
        txtTestCase.KeyPress += TxtTestCase_KeyPress;
        // 
        // FormAddProduct_Test
        // 
        AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
        AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
        ClientSize = new System.Drawing.Size(600, 300);
        Controls.Add(lblTitle);
        Controls.Add(lblTestCase);
        Controls.Add(txtTestCase);
        Controls.Add(btnRun);
        Text = "Chương trình test chức năng thêm sản phẩm";
        ResumeLayout(false);
        PerformLayout();
    }
    #endregion
}