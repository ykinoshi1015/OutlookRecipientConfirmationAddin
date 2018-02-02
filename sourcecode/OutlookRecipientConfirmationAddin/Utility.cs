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
        public enum SendType { Mail, Meeting, Appointment, MeetingResp };

        private const string TANTOU = "担当";
        private const string PROPTAG_URL = "http://schemas.microsoft.com/mapi/proptag/0x0C190102";

        public static string Formatting(RecipientInformationDto recipientInformation)
        {
            string formattedRecipient;

            /// 名前を表示するとき
            if (!recipientInformation.fullName.Equals(""))
            {
                /// Exchangeアドレス帳で受信者の情報が見つかったとき
                if (recipientInformation.division != null)
                {
                    formattedRecipient = string.Format("{0} {1} ({2}【{3}】)", recipientInformation.fullName, recipientInformation.jobTitle, recipientInformation.division, recipientInformation.companyName);
                }
                /// グループ名のみを表示するとき
                else
                {
                    formattedRecipient = recipientInformation.fullName;
                }
            }
            /// 受信者の情報が見つからなかったとき、例外のとき
            else
            {
                /// アドレスだけ表示する
                formattedRecipient = recipientInformation.emailAddress;
            }

            return formattedRecipient;
        }


        /// <summary>
        /// アイテムから、Recipientのリスト取得する
        /// </summary>
        /// <param name="item"></param>
        /// <returns>Recipientsインスタンス(nullの場合メールでも会議でもない)</returns>
        public List<Outlook.Recipient> getRecipients(object Item, ref SendType type, bool IgnoreMeetingResponse)
        {
            Outlook.Recipients recipients = null;

            Outlook.MailItem mail = Item as Outlook.MailItem;
            /// MailItemの場合
            if (mail != null)
            {
                recipients = mail.Recipients;
                type = SendType.Mail;
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
                    /// 会議招集の返信をする場合、宛先確認画面が表示されないようnullを返す
                    if (IgnoreMeetingResponse)
                    {
                        return null;
                    }
                    /// 会議招集の返信の宛先リストを見る場合
                    type = SendType.MeetingResp;
                }
                else
                {
                    //会議招集依頼を送信する場合など
                    type = SendType.Meeting;
                }
                recipients = meeting.Recipients;
            }
            else if (Item is Outlook.AppointmentItem)
            {
                Outlook.AppointmentItem appointment = Item as Outlook.AppointmentItem;
                recipients = appointment.Recipients;
                type = SendType.Appointment;
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
        /// <returns></returns>
        public RecipientInformationDto GetSenderInfomation(object Item)
        {
            Outlook.AddressEntry sender = null;
            Outlook.ExchangeUser exchUser = null;
            Outlook.PropertyAccessor propAccess = null;
            RecipientInformationDto senderInformation = null;

            if (Item is Outlook.MailItem)
            {
                Outlook.MailItem mail = (Item as Outlook.MailItem);

                ///送信元のアカウントのユーザーに対応するSenderプロパティを取得
                sender = mail.Sender;
                if (sender != null)
                {
                    Outlook.Recipient recResolve = Globals.ThisAddIn.Application.Session.CreateRecipient(sender.Address);
                    exchUser = recResolve.AddressEntry.GetExchangeUser();
                }
                /// 新規メッセージ編集中/送信時はSenderはnullなので、SenderEmailAddressからExchangeUserを探す
                else if (mail.SenderEmailAddress != null)
                {
                    Outlook.Recipient recResolve = Globals.ThisAddIn.Application.Session.CreateRecipient(mail.SenderEmailAddress);
                    exchUser = recResolve.AddressEntry.GetExchangeUser();
                }
            }
            /// MeetingItemの場合
            else if (Item is Outlook.MeetingItem)
            {
                Outlook.MeetingItem meeting = Item as Outlook.MeetingItem;

                //送信者のExchangeUserを取得
                propAccess = meeting.PropertyAccessor;
                exchUser = FindExchangeUser(propAccess, ref sender);
            }
            /// AppointmentItemの場合(送信前の会議系のメール)
            else if (Item is Outlook.AppointmentItem)
            {
                Outlook.AppointmentItem appointment = Item as Outlook.AppointmentItem;

                //送信者のExchangeUserを取得
                propAccess = appointment.PropertyAccessor;
                exchUser = FindExchangeUser(propAccess, ref sender);

            }

            if (exchUser != null)
            {
                senderInformation = SetSenderInformation(exchUser);
            }

            /// ExchangeUserが取得できないが、送信者はいる場合
            else if (sender != null)
            {
                senderInformation = new RecipientInformationDto(sender.Name, Outlook.OlMailRecipientType.olOriginator);
            }

            return senderInformation;
        }

        /// <summary>
        /// 送信者のExchangeUserを取得する（MeetingItemとAppointmentItem用）        
        /// <param name="propAccess"></param>
        /// <returns></returns>
        /// </summary>
        private Outlook.ExchangeUser FindExchangeUser(Outlook.PropertyAccessor propAccess, ref Outlook.AddressEntry sender)
        {
            /// PropoertyAccessorで、送信者の情報を取得
            string senderID = propAccess.BinaryToString(propAccess.GetProperty(PROPTAG_URL));
            sender = Globals.ThisAddIn.Application.Session.GetAddressEntryFromID(senderID);

            /// 送信者のメールアドレスから、ExchangeUserを見つける
            Outlook.Recipient recResolve = Globals.ThisAddIn.Application.Session.CreateRecipient(sender.Address);
            Outlook.ExchangeUser exchUser = recResolve.AddressEntry.GetExchangeUser();

            return exchUser;
        }

        /// <summary>
        /// 送信者のExchangeUserプロパティから、表示に必要な情報をセットする
        /// </summary>
        /// <param name="exchUser"></param>
        /// <returns></returns>
        private RecipientInformationDto SetSenderInformation(Outlook.ExchangeUser exchUser)
        {
            Outlook.ContactItem contactItem = null;

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

            return new RecipientInformationDto(contactItem.FullName, contactItem.Department,
                contactItem.CompanyName, jobTitle, Outlook.OlMailRecipientType.olOriginator);

        }

    }
}
