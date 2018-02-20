using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Drawing;
using System.Windows.Forms;
using static OutlookRecipientConfirmationAddin.RecipientConfirmationWindow;
using Office = Microsoft.Office.Core;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookRecipientConfirmationAddin
{
    [ComVisible(true)]
    public class RecipientListRibbon : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;

        public RecipientListRibbon()
        {
        }

        #region IRibbonExtensibility のメンバー

        /// <summary>
        /// リボンを定義したXMLファイルを取得する
        /// </summary>
        /// <param name="ribbonID"></param>
        /// <returns></returns>
        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("OutlookRecipientConfirmationAddin.RecipientListRibbon.xml");
        }

        #endregion

        #region リボンのコールバック

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;

        }

        public Bitmap LoadImage(string imageName)
        {
            var assembly = Assembly.GetExecutingAssembly();
            var stream = assembly.GetManifestResourceStream("OutlookRecipientConfirmationAddin." + imageName);

            return new Bitmap(stream);
        }

        #region 宛先確認機能
        /// <summary>
        /// リボンの「宛先確認」ボタンが押された場合の処理
        /// </summary>
        /// <param name="ribbonUI"></param>
        public void RecipientListButton_Click(Office.IRibbonControl ribbonUI)
        {
            try
            {
                FindSelectedItem();
            }
            //例外が発生した場合、エラーダイアログを表示
            catch (Exception ex)
            {
                // エラーダイアログの呼び出し
                ErrorDialog.ShowException(ex);
            }
        }

        /// <summary>
        /// 表示されているアイテムの取得
        /// </summary>
        private void FindSelectedItem()
        {
            /// 一番上にあるエクスプローラorインスペクターを取得
            object activeItem;
            activeItem = Globals.ThisAddIn.Application.ActiveWindow();

            /// エクスプローラ（クリックで選択）の場合
            if (activeItem is Outlook.Explorer)
            {
                activeItem = Globals.ThisAddIn.Application.ActiveExplorer().Selection[1];
            }
            /// インスペクター（ダブルクリックで選択）の場合
            else if (activeItem is Outlook.Inspector)
            {
                activeItem = Globals.ThisAddIn.Application.ActiveInspector().CurrentItem;
            }

            ShowRecipientListWindow(activeItem);
        }

        /// <summary>
        /// 送信者、To、Cc、Bccを取得と検索し、宛先リスト画面を呼び出す
        /// </summary>
        private void ShowRecipientListWindow(object activeItem)
        {
            /// Mailで初期化
            Utility.OutlookItemType itemType = Utility.OutlookItemType.Mail;
            RecipientInformationDto senderInformation = null;

            /// メールの宛先を取得
            List<Outlook.Recipient> recipientsList = new List<Outlook.Recipient>();
            recipientsList = Utility.GetRecipients(activeItem, ref itemType);

            ///  会議招集の返信の場合、そのまま送信する
            if (recipientsList == null)
            {
                return;
            }

            /// 検索し、受信者の宛先情報リストが戻ってくる
            SearchRecipient searchRecipient = new SearchRecipient();
            List<RecipientInformationDto> recipientList = searchRecipient.SearchContact(recipientsList);

            /// 送信者のExchangeUserオブジェクトを取得
            senderInformation = Utility.GetSenderInfomation(activeItem);

            /// 受信者の宛先情報のリストに、送信者の情報も追加する
            if (senderInformation != null)
            {
                recipientList.Add(senderInformation);
            }

            // 宛先リストの画面を表示する
            RecipientListWindow recipientListWindow = new RecipientListWindow(itemType, recipientList);
            recipientListWindow.ShowDialog();
        }
        #endregion

        //#region 添付ファイル付き返信機能
        ///// <summary>
        ///// 添付ファイル付き返信ボタンのクリックイベント
        ///// </summary>
        //public void ReplayWithAttachments_Click(Office.IRibbonControl ribbonUI)
        //{
        //    Outlook.Explorer explorer = Globals.ThisAddIn.Application.ActiveExplorer();
        //    foreach (object selectedItem in explorer.Selection)
        //    {
        //        CreateReplyAllwithAttachment(selectedItem, false);
        //    }
        //}

        ///// <summary>
        ///// 添付ファイル付き全員に返信ボタンのクリックイベント
        ///// </summary>
        //public void ReplayAllWithAttachments_Click(Office.IRibbonControl ribbonUI)
        //{
        //    Outlook.Explorer explorer = Globals.ThisAddIn.Application.ActiveExplorer();
        //    foreach (object selectedItem in explorer.Selection)
        //    {
        //        CreateReplyAllwithAttachment(selectedItem, true);
        //    }
        //}

        ///// <summary>
        ///// 添付ファイル付き返信メールを作成し、表示する
        ///// </summary>
        ///// <param name="targetItem">MailItem or MeetingItemのオブジェクト</param>
        ///// <param name="replyAll">trueなら全員に返信する</param>
        //private void CreateReplyAllwithAttachment(object targetItem, bool replyAll)
        //{
        //    try
        //    {
        //        ///メール
        //        if (targetItem is Outlook.MailItem)
        //        {
        //            Outlook.MailItem mailItem = targetItem as Outlook.MailItem;
        //            Outlook.MailItem replymailitem = (Outlook.MailItem)Globals.ThisAddIn.Application.CreateItem(Microsoft.Office.Interop.Outlook.OlItemType.olMailItem);

        //            replymailitem = mailItem.Forward(); //Create a object as that of Forward as it automatically includes attachments as well

        //            if (replyAll)
        //            {
        //                replymailitem.To = mailItem.To;
        //                replymailitem.CC = mailItem.CC;
        //            }
        //            else
        //            {
        //                replymailitem.To = mailItem.SenderName;
        //            }
        //            replymailitem.Recipients.ResolveAll();
        //            replymailitem.Subject = CreateReplySubject(mailItem.Subject); //same subject +'RE:'              

        //            replymailitem.Display(false);
        //        }
        //        ///会議招待
        //        else if (targetItem is Outlook.MeetingItem)
        //        {
        //            Outlook.MeetingItem meetingItem = targetItem as Outlook.MeetingItem;
        //            Outlook.MailItem replymailitem;
        //            if (replyAll)
        //                replymailitem = meetingItem.ReplyAll();
        //            else
        //                replymailitem = meetingItem.Reply();

        //            /// 受信したMeetingItemに添付されているファイルをいったん別ファイルに保存し、
        //            /// それを返信用MailItemに添付する。
        //            List<string> tmpFiles = new List<string>();
        //            foreach (Outlook.Attachment attachment in meetingItem.Attachments)
        //            {
        //                string tmpFile = Path.GetTempPath() + attachment.FileName;
        //                tmpFiles.Add(tmpFile);
        //                attachment.SaveAsFile(tmpFile);
        //                replymailitem.Attachments.Add(tmpFile);
        //            }
        //            replymailitem.Display(false);

        //            ///別ファイルに保存した添付ファイルはもう不要のため削除する
        //            foreach (string tmpFile in tmpFiles)
        //            {
        //                File.Delete(tmpFile);
        //            }
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBox.Show(ex.Message, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
        //    }
        //}

        ///// <summary>
        ///// 返信用の件名を作成する
        ///// </summary>
        ///// <param name="originalSubject">返信元の件名</param>
        ///// <returns>先頭に"RE:"が付いた件名</returns>
        //private string CreateReplySubject(string originalSubject)
        //{
        //    string mailSubject = String.Empty;
        //    if (originalSubject != null)
        //        mailSubject = originalSubject.Trim();

        //    if (mailSubject.StartsWith("RE:"))
        //    {
        //        mailSubject = mailSubject.Remove(0, 3);
        //        mailSubject = mailSubject.Trim();
        //    }

        //    if (mailSubject.StartsWith("FW:"))
        //    {
        //        mailSubject = mailSubject.Remove(0, 3);
        //        mailSubject = mailSubject.Trim();
        //    }

        //    return "RE: " + mailSubject;
        //}
        //#endregion

        #endregion

        #region ヘルパー

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }

        #endregion
    }
}
