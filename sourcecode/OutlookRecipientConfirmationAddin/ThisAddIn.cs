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
        public Outlook.Inspectors inspectors;

        /// <summary>
        /// アドインが読み込まれると実行される
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            /// 送信イベントの時
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
            try
            {
                Utility.OutlookItemType itemType = Utility.OutlookItemType.Mail;
                RecipientInformationDto senderInformation = null;
                 
                /// メールの宛先を取得
                List<Outlook.Recipient> recipientsList = new List<Outlook.Recipient>();
                recipientsList = Utility.getRecipients(Item, ref itemType, true);

                /// 会議の招待に対する返事の場合
                if (itemType == Utility.OutlookItemType.MeetingResponse)
                {
                    return;
                }

                /// 引数にRecipientのリストを渡すと、宛先情報のリストが戻ってくる
                SearchRecipient searchRecipient = new SearchRecipient();
                List<RecipientInformationDto> recipientList = searchRecipient.SearchContact(recipientsList);

                /// 送信者のExchangeUserオブジェクトを取得
                senderInformation = Utility.GetSenderInfomation(Item);
                
                /// 受信者の宛先情報のリストに、送信者の情報も追加する
                if (senderInformation != null)
                {
                    recipientList.Add(senderInformation);
                }

                /// 引数に宛先詳細を渡し、確認フォームを表示する
                RecipientConfirmationWindow recipientConfirmationWindow = new RecipientConfirmationWindow(itemType, recipientList);
                DialogResult result = recipientConfirmationWindow.ShowDialog();

                /// 画面でOK以外が選択された場合
                if (result != DialogResult.OK)
                {
                    //メール送信のイベントをキャンセルする
                    Cancel = true;
                }
            }
            /// 何らかのエラーが発生したらイベントをキャンセルする
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                Cancel = true;
            }
        }

        /// <summary>
        /// リボン (XML) アイテムを有効にする
        /// </summary>
        /// <returns></returns>
        protected override Office.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new RecipientListRibbon();
        }
    }
}
