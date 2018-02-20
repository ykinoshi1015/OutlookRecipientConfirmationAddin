using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookRecipientConfirmationAddin
{
    /// <summary>
    /// 送信時の宛先表示とリボンの共通処理を入れるクラス
    /// </summary>
    public class Utility
    {
        /// アイテムの種類
        public enum OutlookItemType { Mail, Meeting, Appointment, MeetingResponse };

        private const string TANTOU = "担当";

        /// <summary>
        /// アイテムから、宛先(Recipient)のリスト取得する
        /// </summary>
        /// <param name="Item">Outlookアイテムオブジェクト</param>
        /// <param name="type">アイテムの種類</param>
        /// <param name="IgnoreMeetingResponse">会議招集の返信かどうか</param>
        /// <returns>Recipientsインスタンス(会議招集の返信や、MailItem,MeetingItem,AppointmentItemでない場合null)</returns>
        public static List<Outlook.Recipient> GetRecipients(object Item, ref OutlookItemType type, bool IgnoreMeetingResponse = false)
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

                /// 会議招集の返信の場合
                /// "IPM.Schedule.Meeting.Resp.Neg";
                /// "IPM.Schedule.Meeting.Resp.Pos";
                /// "IPM.Schedule.Meeting.Resp.Tent";
                if (meeting.MessageClass.Contains("IPM.Schedule.Meeting.Resp."))
                {
                    type = OutlookItemType.MeetingResponse;

                    /// 会議招集の返信をする場合、宛先確認画面が表示されないようnullを返す
                    if (IgnoreMeetingResponse)
                    {
                        return null;
                    }
                }
                /// 会議出席依頼を送信する場合など
                else
                {
                    type = OutlookItemType.Meeting;
                }
                recipients = meeting.Recipients;
            }
            /// AppointmentItemの場合（編集中の会議招集メール）
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
        /// <param name="Item">Outlookアイテムオブジェクト</param>
        /// <returns>送信者の宛先情報インスタンス（送信者が取得できない場合null）</returns>
        public static RecipientInformationDto GetSenderInfomation(object Item)
        {
            Outlook.AddressEntry sender = null;
            Outlook.ExchangeUser exchUser = null;
            RecipientInformationDto senderInformation = null;
            Outlook.Recipient recResolve;

            /// MailItemの場合
            if (Item is Outlook.MailItem)
            {
                Outlook.MailItem mail = (Item as Outlook.MailItem);

                ///送信元のアカウントに対応するSenderプロパティを取得
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

                /// SenderEmailAddressから、送信者のAddressEntry及びExchangeUserを取得
                recResolve = Globals.ThisAddIn.Application.Session.CreateRecipient(meeting.SenderEmailAddress);
                sender = recResolve.AddressEntry;
                exchUser = sender.GetExchangeUser();
            }
            /// AppointmentItemの場合
            else if (Item is Outlook.AppointmentItem)
            {
                Outlook.AppointmentItem appointment = Item as Outlook.AppointmentItem;

                /// 編集中のユーザを送信者として取得
                sender = Globals.ThisAddIn.Application.Session.CurrentUser.AddressEntry;
                exchUser = sender.GetExchangeUser();
            }

            /// 送信者のExchangeUserが取得できた場合
            if (exchUser != null)
            {
                senderInformation = new RecipientInformationDto(exchUser.Name,
                                                                exchUser.Department,
                                                                exchUser.CompanyName,
                                                                FormatJobTitle(exchUser.JobTitle),
                                                                Outlook.OlMailRecipientType.olOriginator); 
            }
            /// ExchangeUserが取得できないが、送信者はいる場合
            else if (sender != null)
            {
                senderInformation = new RecipientInformationDto(sender.Name, Outlook.OlMailRecipientType.olOriginator);
            }

            return senderInformation;
        }

        /// <summary>
        /// 表示する必要のない役職の場合、空文字を入れる
        /// </summary>
        /// <param name="jobTitle">ExchangeUser の役職</param>
        /// <returns>表示用の役職</returns>
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
