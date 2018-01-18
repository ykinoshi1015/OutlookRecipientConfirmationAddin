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

    public partial class RecipientListWindow : Form
    {
        RecipientConfirmationWindow.SendType _type;
        List<RecipientInformationDto> _recipients;

        public RecipientListWindow()
        {
            InitializeComponent();
        }

        public RecipientListWindow(RecipientConfirmationWindow.SendType type, List<RecipientInformationDto> recipients)
        {
            InitializeComponent();
            this._type = type;
            this._recipients = recipients;
        }

        private void RecipientListWindow_Load(object sender, EventArgs e)
        {
            string firstHeader = "", secondHeder = "", thirdHeader = "";
            switch (_type)
            {
                case RecipientConfirmationWindow.SendType.Mail:
                    firstHeader = "To";
                    secondHeder = "Cc";
                    thirdHeader = "Bcc";
                    break;

                case RecipientConfirmationWindow.SendType.Meeting:
                    firstHeader = "参加者";
                    secondHeder = "参加者(任意)";
                    thirdHeader = "リソース";
                    break;
            }

            textBox1.Text = string.Format("■------------ {0}: {1}件 ------------■\r\n", firstHeader, 12345);
            foreach (var recipients in _recipients)
            {
                textBox1.Text += recipients + "\r\n";
            }
            textBox1.AppendText("\r\n");

            //textBox1.Text += string.Format("■------------ {0}: {1}件 ------------■\r\n", secondHeder, ccList.Count());
            //foreach (var recipients in ccList)
            //{
            //    textBox1.Text += recipients + "\r\n";
            //}
            //textBox1.AppendText("\r\n");

            //textBox1.Text += string.Format("■------------ {0}: {1}件 ------------■\r\n", thirdHeader, bccList.Count());
            //foreach (var recipients in bccList)
            //{
            //    textBox1.Text += recipients + "\r\n";
            //}

            /// 読み取り専用、自動で折り返さない
            textBox1.ReadOnly = true;
            textBox1.WordWrap = false;

            /// 必要な場合、垂直、水平両方のスクロールバーを表示
            textBox1.ScrollBars = ScrollBars.Both;

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
    }
}
