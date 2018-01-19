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
        List<RecipientInformationDto> _recipientsList;

        public RecipientListWindow()
        {
            InitializeComponent();
        }

        public RecipientListWindow(RecipientConfirmationWindow.SendType type, List<RecipientInformationDto> recipients)
        {
            InitializeComponent();
            this._type = type;
            this._recipientsList = recipients;
        }

        private void RecipientListWindow_Load(object sender, EventArgs e)
        {
            /// 表示用にフォーマッティングした宛先と、その宛先を入れるリスト
            string formattedRecipient;
            List<string> toList = new List<string>();
            List<string> ccList = new List<string>();
            List<string> bccList = new List<string>();


            /// 宛先をフォーマッティングする
            foreach (var recipient in _recipientsList)
            {
                /// 宛先の何を表示するか
                /// 名前を表示するとき
                if (!recipient.fullName.Equals(""))
                {
                    /// Exchangeアドレス帳で受信者の情報が見つかったとき
                    if (recipient.division != null)
                    {
                        formattedRecipient = string.Format("{0} {1} ({2}【{3}】)", recipient.fullName, recipient.jobTitle, recipient.division, recipient.companyName);
                    }
                    /// グループ名のみを表示するとき
                    else
                    {
                        formattedRecipient = recipient.fullName;
                    }
                }
                /// 受信者の情報が見つからなかったとき、例外のとき
                else
                {
                    /// アドレスだけ表示する
                    formattedRecipient = recipient.emailAddress;
                }

                /// 宛先タイプごとにリストに追加
                switch (recipient.recipientType)
                {
                    case Outlook.OlMailRecipientType.olTo:
                        toList.Add(formattedRecipient);
                        break;

                    case Outlook.OlMailRecipientType.olCC:
                        ccList.Add(formattedRecipient);
                        break;

                    case Outlook.OlMailRecipientType.olBCC:
                        bccList.Add(formattedRecipient);
                        break;
                }
            }

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

            textBox1.Text = string.Format("■------------ {0}: {1}件 ------------■\r\n", firstHeader, toList.Count());
            foreach (var recipient in toList)
            {
                textBox1.Text += recipient + "\r\n";
            }
            textBox1.AppendText("\r\n");

            textBox1.Text += string.Format("■------------ {0}: {1}件 ------------■\r\n", secondHeder, ccList.Count());
            foreach (var recipient in ccList)
            {
                textBox1.Text += recipient + "\r\n";
            }
            textBox1.AppendText("\r\n");

            textBox1.Text += string.Format("■------------ {0}: {1}件 ------------■\r\n", thirdHeader, bccList.Count());
            foreach (var recipient in bccList)
            {
                textBox1.Text += recipient + "\r\n";
            }

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

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }
    }
}
