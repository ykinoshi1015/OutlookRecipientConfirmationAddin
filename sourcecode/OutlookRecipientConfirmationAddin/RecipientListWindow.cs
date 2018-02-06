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
        Utility.OutlookItemType _type;
        List<RecipientInformationDto> _recipientsList;

        public RecipientListWindow()
        {
            InitializeComponent();
        }

        public RecipientListWindow(Utility.OutlookItemType type, List<RecipientInformationDto> recipients) : base(type, recipients)
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

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void label2_Click(object sender, EventArgs e)
        {

        }
    }
}
