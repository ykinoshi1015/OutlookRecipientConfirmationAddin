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
        public RecipientConfirmationWindow(Utility.OutlookItemType type, List<RecipientInformationDto> recipients) : base(type, recipients) 
        {
            InitializeComponent();
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

    }
}
