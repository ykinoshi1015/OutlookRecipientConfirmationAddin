﻿using System;
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

            Microsoft.Win32.RegistryKey regkey = null;

            //Outlookのバージョン番号を取得
            string version = Application.Version;

            //Outlook2016を使っている場合
            if (version.StartsWith("16."))
            {
                //アドイン無効化の監視対象の設定をするキーを開く
                //キーが存在しないときは新しく作成
                regkey = Microsoft.Win32.Registry.CurrentUser.CreateSubKey(@"Software\Microsoft\Office\16.0\Outlook\Resiliency\DoNotDisableAddinList");
            }
            //Outlook2013を使っている場合
            else if (version.StartsWith("15."))
            {
                regkey = Microsoft.Win32.Registry.CurrentUser.CreateSubKey(@"Software\Microsoft\Office\15.0\Outlook\Resiliency\DoNotDisableAddinList");
            }

            //REG_DWORDで書き込む
            regkey.SetValue("OutlookRecipientConfirmationAddin", 0x000000001, Microsoft.Win32.RegistryValueKind.DWord);

            //閉じる
            regkey.Close();
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
                RecipientConfirmationWindow.SendType itemType = RecipientConfirmationWindow.SendType.Mail;

                /// メールでも会議招集でもなければ、そのまま送信する
                Outlook.Recipients recipients = getRecipients(Item, ref itemType);
                if (recipients == null)
                {
                    return;
                }

                /// 受信者の情報をリストする
                List<Outlook.Recipient> recipientsList = new List<Outlook.Recipient>();
                foreach (Outlook.Recipient recipient in recipients)
                {
                    recipientsList.Add(recipient);
                }

                /// 検索クラスを呼び出す
                SearchRecipient searchRecipient = new SearchRecipient();

                /// 引数にTO, CC, BCCに入力されたアドレスのリストを渡すと、宛先情報のリストが戻ってくる
                List<RecipientInformationDto> recipientList = searchRecipient.SearchContact(recipientsList);

                /// 表示用にフォーマッティングするクラス
                List<string> formattedToList = new List<string>();
                List<string> formattedCcList = new List<string>();
                List<string> formattedBccList = new List<string>();

                /// 受信者のタイプに応じたリストに、フォーマットしてから追加する
                foreach (var recipientInformation in
                    recipientList)
                {
                    switch (recipientInformation.recipientType)
                    {

                        case Outlook.OlMailRecipientType.olTo:
                            formattedToList.Add(Utility.Formatting(recipientInformation));
                            break;

                        case Outlook.OlMailRecipientType.olCC:
                            formattedCcList.Add(Utility.Formatting(recipientInformation));
                            break;

                        case Outlook.OlMailRecipientType.olBCC:
                            formattedBccList.Add(Utility.Formatting(recipientInformation));
                            break;
                    }

                }

                /// 引数に宛先詳細を渡し、確認フォームを表示する
                RecipientConfirmationWindow recipientConfirmationWindow = new RecipientConfirmationWindow(itemType, formattedToList, formattedCcList, formattedBccList);
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
        /// ItemからMailItem or MettingItemのRecipientsの取得する
        /// </summary>
        /// <param name="item"></param>
        /// <returns>Recipientsインスタンス(nullの場合メールでも会議でもない)</returns>
        private Outlook.Recipients getRecipients(object Item, ref RecipientConfirmationWindow.SendType type)
        {
            Outlook.Recipients recipients = null;


            Outlook.MailItem mail = Item as Outlook.MailItem;
            if (mail != null)
            {
                recipients = mail.Recipients;
                type = RecipientConfirmationWindow.SendType.Mail;
            }
            else
            {
                Outlook.MeetingItem meeting = Item as Outlook.MeetingItem;
                if (meeting != null)
                {
                    if (meeting.MessageClass.Contains("IPM.Schedule.Meeting.Resp."))
                    {
                        //会議招集の返信
                        //"IPM.Schedule.Meeting.Resp.Neg";
                        //"IPM.Schedule.Meeting.Resp.Pos";
                        //"IPM.Schedule.Meeting.Resp.Tent";

                        // 宛先確認画面が表示されないようnullを返す
                        return null;
                    }
                    else
                    {
                        //会議招集依頼など
                        //"IPM.Schedule.Meeting.Request";
                        //"IPM.Schedule.Meeting.Canceled";
                        //"IPM.Schedule.Meeting.Notification.Forward";

                        recipients = meeting.Recipients;
                        type = RecipientConfirmationWindow.SendType.Meeting;
                    }
                }
            }
            return recipients;
        }
    }
}
