using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace OutlookRecipientConfirmationAddin
{
    /// <summary>
    /// エラーダイアログのクラス
    /// </summary>
    public partial class ErrorDialog : Form
    {
        Exception _ex;

        public ErrorDialog()
        {
            InitializeComponent();
        }

        public ErrorDialog(Exception ex)
        {
            InitializeComponent();
            _ex = ex;
        }

        private void ErrorDialog_Load(object sender, EventArgs e)
        {
            FormatTextBox(_ex);
        }

        private void FormatTextBox(Exception _ex)
        {
            textBox1.Text = _ex.Message + "\r\n";
            textBox1.AppendText("\r\n");
            textBox1.Text += _ex.StackTrace + "\r\n";
        }

        private void GitHub_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            System.Diagnostics.Process.Start("https://github.com/ykinoshi1015/OutlookRecipientConfirmationAddin");
        }

    }
}
