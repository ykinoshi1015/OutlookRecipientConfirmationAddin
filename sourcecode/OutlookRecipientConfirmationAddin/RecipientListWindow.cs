using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookRecipientConfirmationAddin
{

    public partial class RecipientListWindow : RecipientCommonWindow
    {
        Utility.SendType _type;
        List<RecipientInformationDto> _recipientsList;

        public RecipientListWindow()
        {
            InitializeComponent();
        }

        public RecipientListWindow(Utility.SendType type, List<RecipientInformationDto> recipients) : base(type, recipients)
        {
            InitializeComponent();
            this._type = type;
            this._recipientsList = recipients;
        }

        /// <summary>
        /// リボンクラスから渡された宛先情報（To+Cc+Bcc+送信者）を表示用にフォーマッティングし、宛先画面をつくる
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void RecipientListWindow_Load(object sender, EventArgs e)
        {
            /// baseクラスでテキストボックスの内容を作る
            RecipientCommonWindow_format();
        }

        private void button1_Click(object sender, EventArgs e)
        {

        }

        /// <summary>
        ///  「GitHub」のリンクが押された場合
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            System.Diagnostics.Process.Start("https://github.com/ykinoshi1015/OutlookRecipientConfirmationAddin");
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }
    }
}
