using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
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
        /// リボンを定義したXMLファイrを取得する
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

        /// <summary>
        /// 宛先確認ボタンが押された場合
        /// この中で、そのメールの受信者の一覧を探してきて、次の画面に渡す？
        /// </summary>
        /// <param name="ribbonUI"></param>
        public void RecipientListButton_Click(Office.IRibbonControl ribbonUI)
        {
            try
            {
                /// 選択されているアイテムを取得
                Outlook.NameSpace objNamespace = Globals.ThisAddIn.Application.GetNamespace("MAPI");
                var selectedItems = Globals.ThisAddIn.Application.ActiveExplorer();

                /// 選択されているアイテムが一個の場合のみ、宛先確認を表示
                if (selectedItems.Selection.Count == 1)
                {
                    var selectedItem = selectedItems.Selection[1];

                    Outlook.Recipients recipients = null;

                    /// とりあえずmailにしてみた
                    RecipientConfirmationWindow.SendType type = RecipientConfirmationWindow.SendType.Mail;

                    /// 表示しているのがMailItemの場合
                    if (selectedItem is Outlook.MailItem)
                    {
                        Outlook.MailItem mail = (selectedItem as Outlook.MailItem);
                        recipients = mail.Recipients;
                        type = RecipientConfirmationWindow.SendType.Mail;
                    }
                    /// MeetingItemの場合
                    else
                    {
                        Outlook.MeetingItem meeting = selectedItem as Outlook.MeetingItem;
                        if (meeting != null)
                        {
                            recipients = meeting.Recipients;
                            type = RecipientConfirmationWindow.SendType.Meeting;
                        }
                    }

                    /////mailでもmeetingでもなければの処理　いる？
                    //if (recipients == null)
                    //{
                    //    return;
                    //}

                    /// 受信者の情報をリストする
                    List<Outlook.Recipient> recipientsList = new List<Outlook.Recipient>();
                    foreach (Outlook.Recipient recipient in recipients)
                    {
                        recipientsList.Add(recipient);
                    }

                    /// 検索クラスで、引数に宛先に指定されたアドレスのリストを渡すと、宛先情報のリストが戻ってくる
                    SearchRecipient searchRecipient = new SearchRecipient();
                    List<RecipientInformationDto> recipientList = searchRecipient.SearchContact(recipientsList);

                    // 宛先リストの画面を表示する
                    RecipientListWindow recipientListWindow = new RecipientListWindow(type, recipientList);
                    recipientListWindow.ShowDialog();
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show("宛先を表示出来ません");
                Console.WriteLine(ex.Message);
            }
        }
        
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
