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
    public partial class RecipientConfirmationWindow : RecipientCommonWindow
    {

        Utility.OutlookItemType _type;
        List<RecipientInformationDto> _recipientsList;

        public RecipientConfirmationWindow()
        {
            InitializeComponent();
        }

        public RecipientConfirmationWindow(Utility.OutlookItemType type, List<RecipientInformationDto> recipients) : base(type, recipients)
        {
            InitializeComponent();
            this._type = type;
            this._recipientsList = recipients;
        }

        /// <summary>
        /// 宛名確認画面をロード、テキストボックスに値を設定する
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void RecipientConfirmationWindow_Load(object sender, EventArgs e)
        {
            /// baseクラスでテキストボックスの内容を作る
            RecipientCommonWindow_format();
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
        
    }
}
