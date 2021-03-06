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
    public partial class RecipientCommonWindow : Form
    {
        /// 表示しているアイテムのタイプ
        protected Utility.OutlookItemType _type;
        /// 宛先情報のリスト
        protected List<RecipientInformationDto> _recipientsList;
        
        private const string RECIPIENT_HEADER = "■------------ {0}: {1}件 ------------■\r\n";

        private Graphics _measureGraphics;
        private System.Drawing.Font _measureFont;

        private RecipientCommonWindow()
        {
            InitializeComponent();
        }

        protected RecipientCommonWindow(Utility.OutlookItemType type, List<RecipientInformationDto> recipients)
        {
            InitializeComponent();
            _type = type;
            _recipientsList = recipients;

            _measureFont = textBox1.Font;
            _measureGraphics = textBox1.CreateGraphics();
        }

        /// <summary>
        /// リボンクラスから渡された宛先情報（To+Cc+Bcc+送信者）を表示用にフォーマッティングし、宛先画面をつくる
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void RecipientCommonWindow_Load(object sender, EventArgs e)
        {
        }

        protected void RecipientCommonWindow_format()
        {
            /// フォーマッティングした宛先と、それを入れるリスト
            string formattedRecipient;
            string originator = null;
            List<string> toList = new List<string>();
            List<string> ccList = new List<string>();
            List<string> bccList = new List<string>();

            float maxWidth = GetNameAndJobTitleMaxWidth(_recipientsList);

            /// 宛先をフォーマッティングする
            foreach (var recipient in _recipientsList)
            {
                if (!recipient.fullName.Equals(""))
                {
                    /// Exchangeアドレス帳で受信者の情報が見つかった場合（名前、所属情報を表示）
                    if (recipient.division != null)
                    {
                        string nameAndJob = string.Format("{0} {1} ", recipient.fullName, recipient.jobTitle);
                        formattedRecipient = string.Format("{0}({1}【{2}】)", PaddingRight(nameAndJob, maxWidth), recipient.division, recipient.companyName);
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
                if (_type == Utility.OutlookItemType.Task)
                {
                    switch (recipient.recipientType)
                    {
                        case Outlook.OlMailRecipientType.olOriginator:
                            originator = formattedRecipient;
                            break;

                        default:
                            /// Taskの場合、すべて宛先として扱う
                            toList.Add(formattedRecipient);
                            break;
                    }
                }
                else
                {
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
            }

            string firstHeader = "", secondHeder = "", thirdHeader = "";
            switch (_type)
            {
                case Utility.OutlookItemType.Mail:
                case Utility.OutlookItemType.MeetingResponse:
                case Utility.OutlookItemType.Sharing:
                    firstHeader = "To";
                    secondHeder = "Cc";
                    thirdHeader = "Bcc";
                    break;

                case Utility.OutlookItemType.Meeting:
                case Utility.OutlookItemType.Appointment:
                    firstHeader = "参加者";
                    secondHeder = "参加者(任意)";
                    thirdHeader = "リソース";
                    break;

                case Utility.OutlookItemType.Report:
                case Utility.OutlookItemType.Task:
                    firstHeader = "宛先";
                    break;
            }

            textBox1.Text += string.Format("□―――――― 送信者 ――――――□\r\n");
            textBox1.Text += originator + "\r\n";
            textBox1.AppendText("\r\n");

            textBox1.Text += string.Format(RECIPIENT_HEADER, firstHeader, toList.Count());
            foreach (var recipient in toList)
            {
                textBox1.Text += recipient + "\r\n";
            }

            // ReportItemとTaskはCC/BCCが設定できないので、宛先だけ表示すればOK
            if (_type != Utility.OutlookItemType.Report)
            {
                textBox1.AppendText("\r\n");
                textBox1.Text += string.Format(RECIPIENT_HEADER, secondHeder, ccList.Count());
                foreach (var recipient in ccList)
                {
                    textBox1.Text += recipient + "\r\n";
                }
                textBox1.AppendText("\r\n");

                textBox1.Text += string.Format(RECIPIENT_HEADER, thirdHeader, bccList.Count());
                foreach (var recipient in bccList)
                {
                    textBox1.Text += recipient + "\r\n";
                }
            }
            /// 読み取り専用、自動で折り返さない
            textBox1.ReadOnly = true;
            textBox1.WordWrap = false;

            /// 必要な場合、垂直、水平両方のスクロールバーを表示
            textBox1.ScrollBars = ScrollBars.Both;

        }

        private float GetNameAndJobTitleMaxWidth(List<RecipientInformationDto> recipients)
        {
            string maxNameJob = "";
            float maxWidth = 0;

            foreach (RecipientInformationDto recipient in recipients)
            {
                string name_job = string.Format("{0} {1} ", recipient.fullName, recipient.jobTitle);
                float width = TextRenderer.MeasureText(_measureGraphics, name_job, _measureFont).Width;
                if (width >= maxWidth)
                {
                    maxNameJob = name_job;
                    maxWidth = width;
                }
            }
            return maxWidth;
        }

        private string PaddingRight(string src, float paddingWidth)
        {
            string paddingString;
            StringBuilder builder = new StringBuilder();
            builder.Append(src);

            while(true)
            {
                paddingString = builder.ToString();
                float width = TextRenderer.MeasureText(_measureGraphics, paddingString, _measureFont).Width;
                if (width < paddingWidth)
                {
                    builder.Append(' ');
                    continue;
                }
                break;
            }
            return paddingString;
        }

        /// <summary>
        ///  「GitHub」のリンクが押された場合
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        public void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            System.Diagnostics.Process.Start("https://github.com/ykinoshi1015/OutlookRecipientConfirmationAddin");
        }

        private void RecipientCommonWindow_FormClosing(object sender, FormClosingEventArgs e)
        {
            _measureGraphics.Dispose();
        }
    }
}
