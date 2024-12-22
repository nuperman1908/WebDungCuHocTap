using System;
using System.Windows.Forms;

namespace TestingProgram
{
    public partial class FormAddProduct_Test : Form
    {
        public FormAddProduct_Test()
        {
            InitializeComponent();
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>


        #endregion

        private void TxtTestCase_KeyPress(object sender, KeyPressEventArgs e)
        {
            // Chỉ cho phép nhập số
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar))
            {
                e.Handled = true;
            }
        }
        private void TxtTestCase_TextChanged(object sender, EventArgs e)
        {
            // Giới hạn nhập tối đa 6 chữ số
            if (txtTestCase.Text.Length > 6)
            {
                txtTestCase.Text = txtTestCase.Text.Substring(0, 6);
                txtTestCase.SelectionStart = txtTestCase.Text.Length;
            }
        }

        private System.Windows.Forms.Label lblTitle;
        private System.Windows.Forms.Label lblTestCase;
        private System.Windows.Forms.TextBox txtTestCase;
        private System.Windows.Forms.Button btnRun;

        private string webUrl;
        int numTestCases;
        private void btnRun_Click(object sender, EventArgs e)
        {
            if (txtTestCase.Text == "" || txtTestCase.Text == null|| int.Parse(txtTestCase.Text) <= 0)
            {
                MessageBox.Show("Hãy nhập số lượng test case!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }            

            try
            {
                Testing.RunAddProductTest(int.Parse(txtTestCase.Text));
            }
            catch (Exception ex)
            {
                MessageBox.Show("Xuất hiện lỗi trong quá trình chạy: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
