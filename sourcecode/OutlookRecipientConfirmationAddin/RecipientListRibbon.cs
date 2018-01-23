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
        private const string TANTOU = "担当";
        private const string PROPTAG_URL = "http://schemas.microsoft.com/mapi/proptag/0x0C190102";

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
            catch (Exception ex)
            {
                MessageBox.Show("宛先を表示出来ません");
                Console.WriteLine(ex.Message);
            }
        }

        /// <summary>
        /// 画面で選択中のアイテムを取得する
        /// </summary>
        private void FindSelectedItem()
        {
            /// 選択されているアイテムを取得
            Outlook.NameSpace objNamespace = Globals.ThisAddIn.Application.GetNamespace("MAPI");
            Outlook.Explorer selectedItems = Globals.ThisAddIn.Application.ActiveExplorer();

            /// 選択が1個の場合、宛先リストを表示するメソッドを呼ぶ
            if (selectedItems.Selection.Count == 1)
            {
                ShowRecipientListWindow(selectedItems);
            }
            //アイテムが2つ以上選択された場合は、メッセージを表示
            else
            {
                MessageBox.Show("アイテムを1つ選択してください");
            }
        }

        /// <summary>
        /// 送信者、To、Cc、Bccを取得と検索し、宛先リスト画面を呼び出す
        /// </summary>
        private void ShowRecipientListWindow(Outlook.Explorer selectedItems)
        {
            var selectedItem = selectedItems.Selection[1];

            Outlook.Recipients recipients = null;
            Outlook.AddressEntry sender = null;
            RecipientInformationDto senderInformation = null;
            Outlook.ExchangeUser exchUser = null;
            Outlook.PropertyAccessor propAccess = null;

            /// Mailで初期化
            RecipientConfirmationWindow.SendType type = RecipientConfirmationWindow.SendType.Mail;
            
            /// 表示しているのがMailItemの場合
            if (selectedItem is Outlook.MailItem)
            {
                Outlook.MailItem mail = (selectedItem as Outlook.MailItem);
                recipients = mail.Recipients;

                ///送信元のアカウントのユーザーに対応するSenderプロパティを取得
                sender = mail.Sender;
                Outlook.Recipient recResolve = Globals.ThisAddIn.Application.Session.CreateRecipient(mail.SenderEmailAddress);
                exchUser = recResolve.AddressEntry.GetExchangeUser();
            }
            /// MeetingItemの場合
            else if (selectedItem is Outlook.MeetingItem)
            {
                Outlook.MeetingItem meeting = selectedItem as Outlook.MeetingItem;
                type = RecipientConfirmationWindow.SendType.Meeting;
                recipients = meeting.Recipients;
                propAccess = meeting.PropertyAccessor;

                exchUser = FindExchangeUser(propAccess);
            }
            /// AppointmentItemの場合(招待された会議のキャンセル通知)
            else if (selectedItem is Outlook.AppointmentItem)
            {
                Outlook.AppointmentItem appointment = selectedItem as Outlook.AppointmentItem;
                type = RecipientConfirmationWindow.SendType.Appointment;
                recipients = appointment.Recipients;
                propAccess = appointment.PropertyAccessor;

                exchUser = FindExchangeUser(propAccess);
            }

            /// メールでも会議招集でもない場合、なにも起きない
            if (recipients == null)
            {
                return;
            }

            Outlook.ContactItem contactItem = null;

            /// 送信者のExchangeUserが見つかった場合
            if (exchUser != null)
            {
                contactItem = Globals.ThisAddIn.Application.CreateItem(Outlook.OlItemType.olContactItem);
                contactItem.FullName = exchUser.Name;
                contactItem.CompanyName = exchUser.CompanyName;
                contactItem.Department = exchUser.Department;

                /// 表示する役職ならDtoに、違えば空文字を入れる
                string jobTitle = exchUser.JobTitle;
                if (TANTOU.Equals(contactItem.JobTitle) || contactItem.JobTitle == null)
                {
                    jobTitle = "";
                }

                senderInformation = new RecipientInformationDto(contactItem.FullName, contactItem.Department,
                    contactItem.CompanyName, jobTitle, Outlook.OlMailRecipientType.olOriginator);
            }
            /// 送信者のExchangeUserが見つからなかった場合、表示名を表示
            else
            {
                senderInformation = new RecipientInformationDto(sender.Name, Outlook.OlMailRecipientType.olOriginator);
            }

            /// 受信者の情報をリストに入れる
            List<Outlook.Recipient> recipientsList = new List<Outlook.Recipient>();
            foreach (Outlook.Recipient recipient in recipients)
            {
                recipientsList.Add(recipient);
            }

            /// 検索し、受信者の宛先情報リストが戻ってくる
            SearchRecipient searchRecipient = new SearchRecipient();
            List<RecipientInformationDto> recipientList = searchRecipient.SearchContact(recipientsList);

            /// 受信者の宛先情報のリストに、送信者の情報も追加する
            recipientList.Add(senderInformation);

            // 宛先リストの画面を表示する
            RecipientListWindow recipientListWindow = new RecipientListWindow(type, recipientList);
            recipientListWindow.ShowDialog();
        }

        /// <summary>
        /// ExchangeUserを取得する（MeetingItemとAppointmentItemで共通）
        /// </summary>
        /// <param name="propAccess"></param>
        /// <returns></returns>
        private Outlook.ExchangeUser FindExchangeUser(Outlook.PropertyAccessor propAccess)
        {
            Outlook.AddressEntry sender = null;

            string senderID = propAccess.BinaryToString(propAccess.GetProperty(PROPTAG_URL));
            sender = Globals.ThisAddIn.Application.Session.GetAddressEntryFromID(senderID);
            return sender.GetExchangeUser();
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
