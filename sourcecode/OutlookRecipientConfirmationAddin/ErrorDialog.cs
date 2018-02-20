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
    /// 例外発生時に表示するエラーダイアログのクラス
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
        
        /// <summary>
        /// テキストボックス内の表示内容やフォーマットを整える
        /// </summary>
        /// <param name="_ex">発生した例外オブジェクト</param>
        private void FormatTextBox(Exception _ex)
        {
            //例外のメッセージを表示
            textBox1.Text = _ex.Message + "\r\n";
            textBox1.AppendText("\r\n");
            //例外が発生した時点のスタックトレースを表示
            textBox1.Text += _ex.StackTrace + "\r\n";

            //読み取り専用、自動で折り返さない
            textBox1.ReadOnly = true;
            textBox1.WordWrap = false;

            //必要な場合、垂直、水平両方のスクロールバーを表示
            textBox1.ScrollBars = ScrollBars.Both;

            //文字列を全選択しない
            textBox1.SelectionStart = 0;
        }

        //GitHubのリンクが押された場合
        private void GitHub_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            System.Diagnostics.Process.Start("https://github.com/ykinoshi1015/OutlookRecipientConfirmationAddin");
        }
    }
}
