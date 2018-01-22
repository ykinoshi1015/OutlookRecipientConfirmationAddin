﻿using System;
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
    public partial class RecipientConfirmationWindow : Form
    {
        public enum SendType { Mail, Meeting, Appointment };

        SendType _type;
        List<String> toList;
        List<String> ccList;
        List<String> bccList;

        public RecipientConfirmationWindow()
        {
            InitializeComponent();
            Anchor = (AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right);
        }

        public RecipientConfirmationWindow(SendType type, List<String> toList, List<String> ccList, List<String> bccList)
        {
            InitializeComponent();

            _type = type;
            this.toList = toList;
            this.ccList = ccList;
            this.bccList = bccList;
        }

        /// <summary>
        /// 宛名確認画面をロード、テキストボックスに値を設定する
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void RecipientConfirmationWindow_Load(object sender, EventArgs e)
        {
            string firstHeader = "", secondHeder = "", thirdHeader = "";
            switch (_type)
            {
                case SendType.Mail:
                    firstHeader = "To";
                    secondHeder = "Cc";
                    thirdHeader = "Bcc";
                    break;

                case SendType.Meeting:
                    firstHeader = "参加者";
                    secondHeder = "参加者(任意)";
                    thirdHeader = "リソース";
                    break;
            }

            textBox1.Text = string.Format("■------------ {0}: {1}件 ------------■\r\n", firstHeader, toList.Count());
            foreach (var recipients in toList)
            {
                textBox1.Text += recipients + "\r\n";
            }
            textBox1.AppendText("\r\n");

            textBox1.Text += string.Format("■------------ {0}: {1}件 ------------■\r\n", secondHeder, ccList.Count());
            foreach (var recipients in ccList)
            {
                textBox1.Text += recipients + "\r\n";
            }
            textBox1.AppendText("\r\n");

            textBox1.Text += string.Format("■------------ {0}: {1}件 ------------■\r\n", thirdHeader, bccList.Count());
            foreach (var recipients in bccList)
            {
                textBox1.Text += recipients + "\r\n";
            }

            /// 読み取り専用、自動で折り返さない
            textBox1.ReadOnly = true;
            textBox1.WordWrap = false;

            /// 必要な場合、垂直、水平両方のスクロールバーを表示
            textBox1.ScrollBars = ScrollBars.Both;

            /// アンカーを設定
            textBox1.Anchor = (AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right);
            OK.Anchor = (AnchorStyles.Bottom | AnchorStyles.Right);
            Cancel.Anchor = (AnchorStyles.Bottom | AnchorStyles.Right);
        }

        /// <summary>
        /// テキストボックス
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }
        /// <summary>
        /// OKボタンが押された場合
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button1_Click(object sender, EventArgs e)
        {

        }

        /// <summary>
        /// Cancelボタンが押された場合
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button2_Click(object sender, EventArgs e)
        {

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        /// <summary>
        ///  「提供元」のリンクが押された場合
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            System.Diagnostics.Process.Start("https://github.com/ykinoshi1015/OutlookRecipientConfirmationAddin"); 
        }
    }
}
