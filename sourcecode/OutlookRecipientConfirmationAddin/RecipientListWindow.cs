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
        public RecipientListWindow(Utility.OutlookItemType type, List<RecipientInformationDto> recipients) : base(type, recipients)
        {
            InitializeComponent();
        }

        /// <summary>
        /// リボンクラスから渡された宛先情報（To+Cc+Bcc+送信者）を表示用にフォーマッティングし、宛先画面をつくる
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void RecipientListWindow_Load(object sender, EventArgs e)
        {
            if (_type == Utility.OutlookItemType.Report)
            {
                label2.Text = "以下の宛先に送信できませんでした。";
            }
            else
            {
                label2.Text = "このメールは以下の連絡先宛てです。";
            }

            /// baseクラスでテキストボックスの内容を作る
            RecipientCommonWindow_format();
        }
    }
}
