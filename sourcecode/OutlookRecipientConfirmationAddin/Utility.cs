using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookRecipientConfirmationAddin
{
    /// <summary>
    /// 宛先情報を表示用にフォーマッティングするクラス
    /// </summary>
    public class Utility
    {
        public enum OutlookItemType { Mail, Meeting, Appointment, MeetingResponse };

        private const string TANTOU = "担当";
        private const string PROPTAG_URL = "http://schemas.microsoft.com/mapi/proptag/0x0C190102";

        /// <summary>
        /// アイテムから、Recipientのリスト取得する
        /// </summary>
        /// <param name="item"></param>
        /// <returns>Recipientsインスタンス(nullの場合メールでも会議でもない)</returns>
        public static List<Outlook.Recipient> getRecipients(object Item, ref OutlookItemType type, bool IgnoreMeetingResponse = false)
        {
            Outlook.Recipients recipients = null;

            Outlook.MailItem mail = Item as Outlook.MailItem;
            /// MailItemの場合
            if (mail != null)
            {
                recipients = mail.Recipients;
                type = OutlookItemType.Mail;
            }
            /// MeetingItemの場合
            else if (Item is Outlook.MeetingItem)
            {
                Outlook.MeetingItem meeting = Item as Outlook.MeetingItem;

                /// 会議招集
               //"IPM.Schedule.Meeting.Resp.Neg";
                //"IPM.Schedule.Meeting.Resp.Pos";
                //"IPM.Schedule.Meeting.Resp.Tent";
                if (meeting.MessageClass.Contains("IPM.Schedule.Meeting.Resp."))
                {
                    type = OutlookItemType.MeetingResponse;

                    /// 会議招集の返信をする場合、宛先確認画面が表示されないようnullを返す
                    if (IgnoreMeetingResponse)
                    {
                        return null;
                    }
                }
                else
                {
                    //会議招集依頼を送信する場合など
                    type = OutlookItemType.Meeting;
                }
                recipients = meeting.Recipients;
            }
            else if (Item is Outlook.AppointmentItem)
            {
                Outlook.AppointmentItem appointment = Item as Outlook.AppointmentItem;
                recipients = appointment.Recipients;
                type = OutlookItemType.Appointment;
            }

            /// 受信者の情報をリストに入れる
            List<Outlook.Recipient> recipientsList = new List<Outlook.Recipient>();
            foreach (Outlook.Recipient recipient in recipients)
            {
                recipientsList.Add(recipient);
            }

            return recipientsList;
        }

        /// <summary>
        /// 送信者の情報(Dto)を取得する
        /// </summary>
        /// <param name="Item"></param>
        /// <returns>送信者の宛先情報インスタンス（送信者が取得できない場合null）</returns>
        public static RecipientInformationDto GetSenderInfomation(object Item)
        {
            Outlook.AddressEntry sender = null;
            Outlook.ExchangeUser exchUser = null;
            Outlook.PropertyAccessor propAccess = null;
            RecipientInformationDto senderInformation = null;
            Outlook.Recipient recResolve;
            //Outlook.AddressEntry addressEntry;

            if (Item is Outlook.MailItem)
            {
                Outlook.MailItem mail = (Item as Outlook.MailItem);

                ///送信元のアカウントのユーザーに対応するSenderプロパティを取得
                sender = mail.Sender;
                if (sender != null)
                {
                    recResolve = Globals.ThisAddIn.Application.Session.CreateRecipient(sender.Address);
                    exchUser = recResolve.AddressEntry.GetExchangeUser();
                }
                /// 新規メッセージ編集中/送信時はSenderはnullなので、SenderEmailAddressからExchangeUserを探す
                else if (mail.SenderEmailAddress != null)
                {
                    recResolve = Globals.ThisAddIn.Application.Session.CreateRecipient(mail.SenderEmailAddress);
                    exchUser = recResolve.AddressEntry.GetExchangeUser();
                }
            }
            /// MeetingItemの場合
            else if (Item is Outlook.MeetingItem)
            {
                Outlook.MeetingItem meeting = Item as Outlook.MeetingItem;

                //送信者のExchangeUserを取得
                propAccess = meeting.PropertyAccessor;

                Outlook.AddressEntry addressEntry = GetSenderAddressEntry(propAccess);
                exchUser = addressEntry.GetExchangeUser();
            }
            /// AppointmentItemの場合(送信前の会議系のメール)
            else if (Item is Outlook.AppointmentItem)
            {
                Outlook.AppointmentItem appointment = Item as Outlook.AppointmentItem;

                //送信者のExchangeUserを取得
                propAccess = appointment.PropertyAccessor;
                Outlook.AddressEntry addressEntry  = GetSenderAddressEntry(propAccess);
                /// 送信者のメールアドレスから、ExchangeUserを見つける
                //recResolve = Globals.ThisAddIn.Application.Session.CreateRecipient(addressEntry.Address);
                //exchUser = recResolve.AddressEntry.GetExchangeUser();
                exchUser = addressEntry.GetExchangeUser();
            }

            if (exchUser != null)
            {
                senderInformation = FormatSenderInformation(exchUser);
            }

            /// ExchangeUserが取得できないが、送信者はいる場合
            else if (sender != null)
            {
                senderInformation = new RecipientInformationDto(sender.Name, Outlook.OlMailRecipientType.olOriginator);
            }

            return senderInformation;
        }

        /// <summary>
        /// 送信者のAddressEntryrを取得する（MeetingItemとAppointmentItem用）        
        /// <param name="propAccess"></param>
        /// <returns>送信者のAddressEntry</returns>
        /// </summary>
        private static Outlook.AddressEntry GetSenderAddressEntry(Outlook.PropertyAccessor propAccess)
        {
            /// PropoertyAccessorで、送信者の情報を取得
            string senderID = (propAccess.GetProperty(PROPTAG_URL));
            string senderID2 = propAccess.BinaryToString(senderID);
            Outlook.AddressEntry addressEntry = Globals.ThisAddIn.Application.Session.GetAddressEntryFromID(senderID);

            return addressEntry;
        }

        /// <summary>
        /// 送信者のExchangeUserプロパティから、表示に必要な情報を取り出す
        /// </summary>
        /// <param name="exchUser">送信者のExchangeUserインスタンス</param>
        /// <returns>送信者の宛先情報インスタンス</returns>
        private static RecipientInformationDto FormatSenderInformation(Outlook.ExchangeUser exchUser)
        {
            Outlook.ContactItem contactItem = null;

            contactItem = Globals.ThisAddIn.Application.CreateItem(Outlook.OlItemType.olContactItem);
            contactItem.FullName = exchUser.Name;
            contactItem.CompanyName = exchUser.CompanyName;
            contactItem.Department = exchUser.Department;
            string jobTitle = FormatJobTitle(exchUser.JobTitle);

            return new RecipientInformationDto(contactItem.FullName, contactItem.Department,
                contactItem.CompanyName, jobTitle, Outlook.OlMailRecipientType.olOriginator);
        }

        /// <summary>
        /// 表示する必要のない役職違えば空文字を入れる
        /// </summary>
        /// <param name="jobTitle"></param>
        /// <returns></returns>
        public static string FormatJobTitle(string jobTitle)
        {
            if (TANTOU.Equals(jobTitle) || jobTitle == null)
            {
                jobTitle = "";
            }
            return jobTitle;
        }

    }
}
