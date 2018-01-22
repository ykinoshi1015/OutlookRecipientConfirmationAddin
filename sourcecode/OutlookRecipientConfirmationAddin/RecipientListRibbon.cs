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
        /// 職種が担当の場合の定数
        private const string TANTOU = "担当";

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
        /// リボンの「宛先確認」ボタンが押された場合
        /// 送信者、To、Cc、Bccを取得と検索し、宛先リスト画面を呼び出す
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
                    Outlook.AddressEntry sender = null;
                    RecipientInformationDto senderInformation = null;
                    Outlook.ExchangeUser exchUser = null;

                    /// Mailで初期化
                    RecipientConfirmationWindow.SendType type = RecipientConfirmationWindow.SendType.Mail;

                    /// 表示しているのがMailItemの場合
                    if (selectedItem is Outlook.MailItem)
                    {
                        Outlook.MailItem mail = (selectedItem as Outlook.MailItem);
                        recipients = mail.Recipients;

                        ///送信元のアカウントのユーザーに対応するSenderプロパティを取得
                        sender = mail.Sender;
                        //try
                        //{
                        //    Outlook.Recipient recResolve = Globals.ThisAddIn.Application.Session.CreateRecipient(sender.Address);
                        //    /// Exchangeアドレス帳に存在するアドレスなら、exchUserが見つかる
                        //    exchUser = recResolve.AddressEntry.GetExchangeUser();
                        //}
                        //catch (NullReferenceException ex)
                        //{
                            Outlook.Recipient recResolve = Globals.ThisAddIn.Application.Session.CreateRecipient(mail.SenderEmailAddress);
                            exchUser = recResolve.AddressEntry.GetExchangeUser();
                        //    Console.Write(ex.Message);
                        //}
                    }
                    else
                    {
                        Outlook.PropertyAccessor propAccess = null;

                        /// MeetingItemの場合
                        Outlook.MeetingItem meeting = selectedItem as Outlook.MeetingItem;
                        if (meeting != null)
                        {
                            type = RecipientConfirmationWindow.SendType.Meeting;
                            recipients = meeting.Recipients;
                            propAccess = meeting.PropertyAccessor;

                            string senderID = propAccess.BinaryToString(propAccess.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x0C190102"));
                            sender = Globals.ThisAddIn.Application.Session.GetAddressEntryFromID(senderID);
                            exchUser = sender.GetExchangeUser();
                        }

                        /// AppointmentItemの場合(招待された会議のキャンセル通知？)
                        Outlook.AppointmentItem appointment = selectedItem as Outlook.AppointmentItem;
                        if (appointment != null)
                        {
                            type = RecipientConfirmationWindow.SendType.Appointment;
                            recipients = appointment.Recipients;
                            propAccess = appointment.PropertyAccessor;

                            string senderID = propAccess.BinaryToString(propAccess.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x0C190102"));
                            sender = Globals.ThisAddIn.Application.Session.GetAddressEntryFromID(senderID);
                            exchUser = sender.GetExchangeUser();
                        }
                    }

                    /// MailItem,MeetingItem,AppointmentItem 共通の処理
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
                    /// 送信者のExchangeUserが見つからなかった場合
                    else
                    {
                        senderInformation = new RecipientInformationDto(sender.Name, Outlook.OlMailRecipientType.olOriginator);
                    }

                    /// メールでも会議招集でもない場合、なにも起きない
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

                    /// 検索し、宛先情報のリストが戻ってくる
                    SearchRecipient searchRecipient = new SearchRecipient();
                    List<RecipientInformationDto> recipientList = searchRecipient.SearchContact(recipientsList);

                    /// 宛先情報のリストに、送信者の情報も追加する
                    recipientList.Add(senderInformation);

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
