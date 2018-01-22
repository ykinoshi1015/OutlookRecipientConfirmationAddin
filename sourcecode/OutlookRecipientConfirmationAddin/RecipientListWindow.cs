﻿using System;
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

        /// <summary>
        /// リボンクラスから渡された宛先情報（To+Cc+Bcc+送信者）を表示用にフォーマッティングし、宛先画面をつくる
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void RecipientListWindow_Load(object sender, EventArgs e)
        {
            /// フォーマッティングした宛先と、それを入れるリスト
            string formattedRecipient;
            string originator = null;
            List<string> toList = new List<string>();
            List<string> ccList = new List<string>();
            List<string> bccList = new List<string>();

            /// 宛先をフォーマッティングする
            foreach (var recipient in _recipientsList)
            {
                if (!recipient.fullName.Equals(""))
                {
                    /// Exchangeアドレス帳で受信者の情報が見つかった場合（名前、所属情報を表示）
                    if (recipient.division != null)
                    {
                        formattedRecipient = string.Format("{0} {1} ({2}【{3}】)", recipient.fullName, recipient.jobTitle, recipient.division, recipient.companyName);
                    }
                    /// グループ名のみを表示
                    else
                    {
                        formattedRecipient = recipient.fullName;
                    }
                }
                /// 受信者の情報が見つからなかった場合、例外の場合
                else
                {
                    /// アドレスだけ表示
                    formattedRecipient = recipient.emailAddress;
                }

                /// 宛先タイプごとにリストに追加
                switch (recipient.recipientType)
                {
                    case Outlook.OlMailRecipientType.olOriginator:
                        originator = formattedRecipient;
                        break;

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
                case RecipientConfirmationWindow.SendType.Appointment:
                    firstHeader = "参加者";
                    secondHeder = "参加者(任意)";
                    thirdHeader = "リソース";
                    break;
            }

            textBox1.Text += string.Format("■□――――― 送信者 ―――――□■\r\n");
            textBox1.Text += originator + "\r\n";
            textBox1.Text += string.Format("―――――――――――――――――\r\n");
            textBox1.AppendText("\r\n");

            textBox1.Text += string.Format("■------------ {0}: {1}件 ------------■\r\n", firstHeader, toList.Count());
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

            textBox1.Text += string.Format("■------------ {0}: {1}件 -----------■\r\n", thirdHeader, bccList.Count());
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
