using System;
using System.Collections.Generic;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using System.Windows.Forms;
using DoNotDisableAddinUpdater;

namespace OutlookRecipientConfirmationAddin
{
    public partial class ThisAddIn
    {
        /// <summary>
        /// アドインが読み込まれると実行される
        /// </summary>
        /// <param name="sender">イベントの発生源</param>
        /// <param name="e">発生したイベントのインスタンス</param>
        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            // レジストリ確認のDLLを呼び出し、アドイン無効化の監視をしないようにする
            bool doNotDisableAddinListUpdaterResult = DoNotDisableAddinListUpdater.UpdateDoNotDisableAddinList("OutlookRecipientConfirmationAddin", true);

            ///Notesリンクを開こうとしたときに表示される警告を抑制するよう設定する
            ///アドインの設定画面が実装された、その中で設定できるようにする
            ///※起動時の設定は暫定
            //DoNotDisableAddinListUpdater.DisableProtocolSecurityPopup("notes:");

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
        ///  アイテム送信時の宛先表示画面を生成する
        /// </summary>
        /// <param name="item">アイテム</param>
        /// <param name="cancel">送信をしないかどうか</param>
        private void ConfirmContact(object Item, ref bool Cancel)
        {
            Cursor.Current = Cursors.WaitCursor;
            try
            {
                // アイテムタイプをMailで初期化
                Utility.OutlookItemType itemType = Utility.OutlookItemType.Mail;

                // アイテムの宛先を取得
                List<Outlook.Recipient> recipientsList = new List<Outlook.Recipient>();
                recipientsList = Utility.GetRecipients(Item, ref itemType, true);

                // 会議の招待に対する返事の場合、宛先表示しない
                if (itemType == Utility.OutlookItemType.MeetingResponse)
                {
                    return;
                }

                // 宛先情報のリストを取得
                SearchRecipient searchRecipient = new SearchRecipient();
                List<RecipientInformationDto> recipientList = searchRecipient.SearchContact(recipientsList);

                // 送信者のExchangeUserオブジェクトを取得
                RecipientInformationDto senderInformation = null;
                senderInformation = Utility.GetSenderInfomation(Item);

                // 受信者の宛先情報のリストに、送信者の情報も追加する
                if (senderInformation != null)
                {
                    recipientList.Add(senderInformation);
                }

                Cursor.Current = Cursors.Default;

                // 引数に宛先情報を渡し、宛先表示画面を表示する
                RecipientConfirmationWindow recipientConfirmationWindow = new RecipientConfirmationWindow(itemType, recipientList);
                DialogResult result = recipientConfirmationWindow.ShowDialog();

                // 画面でOK以外が選択された場合
                if (result != DialogResult.OK)
                {
                    // メール送信のイベントをキャンセルする
                    Cancel = true;
                }
            }
            /// 例外が発生した場合、エラーダイアログを表示
            /// 送信イベントをキャンセルする
            catch (Exception ex)
            {
                // エラーダイアログの呼び出し
                ErrorDialog.ShowException(ex);

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
