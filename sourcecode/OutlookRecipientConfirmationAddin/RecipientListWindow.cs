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

        public RecipientListWindow()
        {
            InitializeComponent();
        }

        public RecipientListWindow(Utility.OutlookItemType type, List<RecipientInformationDto> recipients) : base(type, recipients)
        {
            InitializeComponent();
            _type = type;
            _recipientsList = recipients;
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
    }
}
