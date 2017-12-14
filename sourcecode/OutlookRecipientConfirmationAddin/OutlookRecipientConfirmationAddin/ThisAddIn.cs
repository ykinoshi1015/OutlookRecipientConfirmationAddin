using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using System.Windows.Forms;

namespace OutlookRecipientConfirmationAddin
{
    public partial class ThisAddIn
    {
        /// <summary>
        /// アドインが読み込まれると実行される
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            Application.ItemSend += new Outlook.ApplicationEvents_11_ItemSendEventHandler(ConfirmContact);
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            //注: Outlook はこのイベントを発行しなくなりました。Outlook が
            //    シャットダウンする際に実行が必要なコードがある場合は、http://go.microsoft.com/fwlink/?LinkId=506785 を参照してください。
        }

        #region VSTO で生成されたコード

        /// <summary>
        /// デザイナーのサポートに必要なメソッドです。
        /// このメソッドの内容をコード エディターで変更しないでください。
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion

        /// <summary>
        /// 宛先確認
        /// </summary>
        /// <param name="item"></param>
        /// <param name="cancel"></param>
        public void ConfirmContact(object Item, ref bool Cancel)
        {

            /// TO, CC, BCCに入力されたアドレスのリスト
            List<String> toList = new List<String>();
            List<String> ccList = new List<String>();
            List<String> bccList = new List<String>();

            Outlook.MailItem mail = Item as Outlook.MailItem;

            /// 受信者のメールアドレスをタイプ別にリストに追加する
            foreach (Outlook.Recipient recipient in mail.Recipients)
            {
                if (recipient.Type == 1) {
                    toList.Add(recipient.Address);
                }
                else if (recipient.Type == 2)
                {
                    ccList.Add(recipient.Address);
                }
                else if (recipient.Type == 3)
                {
                    bccList.Add(recipient.Address);
                }
            }

            /// 検索クラスを呼び出す
            SearchRecipient searchRecipient = new SearchRecipient(toList, ccList, bccList);
            searchRecipient.SearchContact();
        }
    }
}
