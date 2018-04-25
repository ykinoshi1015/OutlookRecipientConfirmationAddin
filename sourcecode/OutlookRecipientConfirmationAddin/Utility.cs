using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
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
        // アイテムの種類
        public enum OutlookItemType { Mail, Meeting, Appointment, MeetingResponse, Sharing, Report, Task, TaskRequestResponse };

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
            // MailItemの場合
            if (Item is Outlook.MailItem)
            {
                type = OutlookItemType.Mail;
                Outlook.MailItem mail = Item as Outlook.MailItem;
                return GetRecipientList(mail);
            }
            // MeetingItemの場合
            else if (Item is Outlook.MeetingItem)
            {
                Outlook.MeetingItem meeting = Item as Outlook.MeetingItem;

                // 会議招集の返信の場合
                // "IPM.Schedule.Meeting.Resp.Neg";
                // "IPM.Schedule.Meeting.Resp.Pos";
                // "IPM.Schedule.Meeting.Resp.Tent";
                if (meeting.MessageClass.Contains("IPM.Schedule.Meeting.Resp."))
                {
                    type = OutlookItemType.MeetingResponse;

                    // 会議招集の返信をする場合、宛先確認画面が表示されないようnullを返す
                    if (IgnoreMeetingResponse)
                    {
                        return null;
                    }
                    else
                    {
                        return GetRecipientList(meeting);
                    }
                }
                // 会議出席依頼を送信する場合など
                else
                {
                    type = OutlookItemType.Meeting;
                    return GetRecipientList(meeting);
                }
            }
            // AppointmentItemの場合（編集中/送信されていない状態でトレイにある会議招集メール、開催者が取り消した会議のキャンセル通知（自分承認済み））
            else if (Item is Outlook.AppointmentItem)
            {
                type = OutlookItemType.Appointment;
                Outlook.AppointmentItem appointment = Item as Outlook.AppointmentItem;
                return GetRecipientList(appointment);
            }
            else if (Item is Outlook.SharingItem)
            {
                type = OutlookItemType.Sharing;
                Outlook.SharingItem sharing = Item as Outlook.SharingItem;
                return GetRecipientList(sharing);
            }
            else if (Item is Outlook.ReportItem)
            {
                type = OutlookItemType.Report;
                Outlook.ReportItem report = Item as Outlook.ReportItem;
                return GetRecipientList(report);
            }
            else if (Item is Outlook.TaskItem)
            {
                return null;
                //type = OutlookItemType.Task;
                //Outlook.TaskItem task = Item as Outlook.TaskItem;
                //return GetRecipientList(task);
            }
            else if (Item is Outlook.TaskRequestItem)
            {
                return null;
                //type = OutlookItemType.Task;
                //Outlook.TaskRequestItem taskRequest = Item as Outlook.TaskRequestItem;
                //return GetRecipientList(taskRequest);
            }
            else if (Item is Outlook.TaskRequestAcceptItem)
            {
                return null;
                //type = OutlookItemType.TaskRequestResponse;
                //Outlook.TaskRequestAcceptItem taskRequestAccept = Item as Outlook.TaskRequestAcceptItem;
                //return GetRecipientList(taskRequestAccept);
            }
            else if (Item is Outlook.TaskRequestDeclineItem)
            {
                return null;
                //type = OutlookItemType.TaskRequestResponse;
                //Outlook.TaskRequestDeclineItem taskRequestDecline = Item as Outlook.TaskRequestDeclineItem;
                //return GetRecipientList(taskRequestDecline);
            }

            throw new NotSupportedException("未対応のOutlook機能です。");
        }

        #region 各ItemのGetRecipientList
        /// <summary>
        /// MailItemから、宛先(Recipient)のリスト取得する
        /// </summary>
        /// <param name="Item">MailItemオブジェクト</param>
        /// <returns>List<Outlook.Recipient></returns>
        private static List<Outlook.Recipient> GetRecipientList(Outlook.MailItem item)
        {
            Outlook.Recipients recipients = item.Recipients;

            List<Outlook.Recipient> recipientsList = new List<Outlook.Recipient>();
            for (int i = 1; i <= recipients.Count; i++)
            {
                recipientsList.Add(recipients[i]);
            }
            return recipientsList;
        }

        /// <summary>
        /// MeetingItemから、宛先(Recipient)のリスト取得する
        /// </summary>
        /// <param name="Item">MeetingItemオブジェクト</param>
        /// <returns>List<Outlook.Recipient></returns>
        private static List<Outlook.Recipient> GetRecipientList(Outlook.MeetingItem item)
        {
            Outlook.Recipients recipients = item.Recipients;
            List<Outlook.Recipient> recipientsList = new List<Outlook.Recipient>();

            for (int i = 1; i <= recipients.Count; i++)
            {
                if (recipients[i].Type == (int)Outlook.OlMeetingRecipientType.olResource)
                {
                    // 拠点の空きリソースから選択したときに、拠点全部のリソースが宛先表示される現象の対策
                    if (recipients[i].Sendable)
                    {
                        recipientsList.Add(recipients[i]);
                    }
                }
                else
                {
                    recipientsList.Add(recipients[i]);
                }
            }
            return recipientsList;
        }

        /// <summary>
        /// AppointmentItemから、宛先(Recipient)のリスト取得する
        /// </summary>
        /// <param name="Item">AppointmentItemオブジェクト</param>
        /// <returns>List<Outlook.Recipient></returns>
        private static List<Outlook.Recipient> GetRecipientList(Outlook.AppointmentItem item)
        {
            Outlook.Recipients recipients = item.Recipients;
            List<Outlook.Recipient> recipientsList = new List<Outlook.Recipient>();

            //AppointmentItemの場合1番目には送信者、2番目から宛先が入っている
            for (int i = 2; i <= recipients.Count; i++)
            {
                recipientsList.Add(recipients[i]);
            }
            return recipientsList;
        }

        /// <summary>
        /// SharingItemから、宛先(Recipient)のリスト取得する
        /// </summary>
        /// <param name="Item">SharingItemオブジェクト</param>
        /// <returns>List<Outlook.Recipient></returns>
        private static List<Outlook.Recipient> GetRecipientList(Outlook.SharingItem item)
        {
            Outlook.Recipients recipients = item.Recipients;
            List<Outlook.Recipient> recipientsList = new List<Outlook.Recipient>();

            //AppointmentItemの場合1番目には送信者、2番目から宛先が入っている
            for (int i = 1; i <= recipients.Count; i++)
            {
                recipientsList.Add(recipients[i]);
            }
            return recipientsList;
        }

        /// <summary>
        /// ReportItemから、宛先(Recipient)のリスト取得する
        /// </summary>
        /// <param name="Item">ReportItemオブジェクト</param>
        /// <returns>List<Outlook.Recipient></returns>
        private static List<Outlook.Recipient> GetRecipientList(Outlook.ReportItem item)
        {
            //ReportItemのままだと送信先が取れないため、
            //いったんIPM.Noteとして別名保存⇒ロードしてからRecipientsを取得する
            Outlook.ReportItem copiedReport = item.Copy();
            copiedReport.MessageClass = "IPM.Note";
            copiedReport.Save();

            //IPM.Noteとして保存してからロードするとMailItemとして扱えるようになる
            var newReportItem = Globals.ThisAddIn.Application.Session.GetItemFromID(copiedReport.EntryID);
            Outlook.MailItem newMailItem = newReportItem as Outlook.MailItem;

            List<Outlook.Recipient> recipientsList = GetRecipientList(newMailItem);

            copiedReport.Delete();
            return recipientsList;
        }

        /// <summary>
        /// TaskItemから、宛先(Recipient)のリスト取得する
        /// </summary>
        /// <param name="Item">TaskItemオブジェクト</param>
        /// <returns>List<Outlook.Recipient></returns>
        private static List<Outlook.Recipient> GetRecipientList(Outlook.TaskItem item)
        {
            Outlook.Recipients recipients = item.Recipients;
            List<Outlook.Recipient> recipientsList = new List<Outlook.Recipient>();

            if (IsSendTaskRequest(item))
            {
                //これから送信するTaskRequestItem
                for (int i = 1; i <= recipients.Count; i++)
                {
                    recipientsList.Add(recipients[i]);
                }
            }
            else
            {
                //受信したTaskRequestItem
                if (item.Owner == null)
                {
                    recipientsList.Add(item.Recipients[1]);
                }
                else
                {
                    Outlook.Recipient ownerRecipient = Globals.ThisAddIn.Application.Session.CreateRecipient(item.Owner);
                    recipientsList.Add(ownerRecipient);
                }
            }
            return recipientsList;
        }

        /// <summary>
        /// TaskRequestItemから、宛先(Recipient)のリスト取得する
        /// </summary>
        /// <param name="Item">TaskRequestItemオブジェクト</param>
        /// <returns>List<Outlook.Recipient></returns>
        private static List<Outlook.Recipient> GetRecipientList(Outlook.TaskRequestItem item)
        {
            Outlook.TaskItem task = item.GetAssociatedTask(false);
            Outlook.Recipients recipients = task.Recipients;
            List<Outlook.Recipient> recipientsList = new List<Outlook.Recipient>();

            if (recipients.Count == 0)
            {
                //ReportItemのままだと送信先が取れないため、
                //いったんIPM.Noteとして別名保存⇒ロードしてからRecipientsを取得する
                Outlook.TaskRequestItem copiedItem = item.Copy();
                copiedItem.MessageClass = "IPM.Note";
                copiedItem.Save();

                //IPM.Noteとして保存してからロードするとMailItemとして扱えるようになる
                var newReportItem = Globals.ThisAddIn.Application.Session.GetItemFromID(copiedItem.EntryID);
                Outlook.MailItem newMailItem = newReportItem as Outlook.MailItem;

                recipientsList = GetRecipientList(newMailItem);

                copiedItem.Delete();
                return recipientsList;
            }

            if (IsSendTaskRequest(task))
            {
                //これから送信するTaskRequestItem
                for (int i = 1; i <= recipients.Count; i++)
                {
                    recipientsList.Add(recipients[i]);
                }
            }
            else
            {
                //受信したTaskRequestItem
                Outlook.Recipient ownerRecipient = Globals.ThisAddIn.Application.Session.CreateRecipient(task.Owner);
                recipientsList.Add(ownerRecipient);
            }
            return recipientsList;
        }

        /// <summary>
        /// TaskRequestAcceptItemから、宛先(Recipient)のリスト取得する
        /// </summary>
        /// <param name="Item">TaskRequestAcceptItemオブジェクト</param>
        /// <returns>List<Outlook.Recipient></returns>
        private static List<Outlook.Recipient> GetRecipientList(Outlook.TaskRequestAcceptItem item)
        {
            //相手からの承認メールなので、必ず自分宛という想定
            List<Outlook.Recipient> recipientList = new List<Outlook.Recipient>();
            recipientList.Add(Globals.ThisAddIn.Application.Session.CurrentUser);
            return recipientList;
        }

        /// <summary>
        /// TaskRequestDeclineItemから、宛先(Recipient)のリスト取得する
        /// </summary>
        /// <param name="Item">TaskRequestDeclineItemオブジェクト</param>
        /// <returns>List<Outlook.Recipient></returns>
        private static List<Outlook.Recipient> GetRecipientList(Outlook.TaskRequestDeclineItem item)
        {
            //相手からの辞退メールなので、必ず自分宛という想定
            List<Outlook.Recipient> recipientList = new List<Outlook.Recipient>();
            recipientList.Add(Globals.ThisAddIn.Application.Session.CurrentUser);
            return recipientList;
        }
        #endregion

        /// <summary>
        /// 送信者の情報(Dto)を取得する
        /// </summary>
        /// <param name="Item">Outlookアイテムオブジェクト</param>
        /// <returns>送信者の宛先情報インスタンス（送信者が取得できない場合null）</returns>
        public static RecipientInformationDto GetSenderInfomation(object Item)
        {
            if (Item is Outlook.MailItem)
            {
                Outlook.MailItem mail = (Item as Outlook.MailItem);
                return GetSenderInformation(mail);
            }
            else if (Item is Outlook.MeetingItem)
            {
                Outlook.MeetingItem meeting = Item as Outlook.MeetingItem;
                return GetSenderInformation(meeting);
            }
            else if (Item is Outlook.AppointmentItem)
            {
                Outlook.AppointmentItem appointment = Item as Outlook.AppointmentItem;
                return GetSenderInformation(appointment);
            }
            else if (Item is Outlook.ReportItem)
            {
                Outlook.ReportItem report = Item as Outlook.ReportItem;
                return GetSenderInformation(report);
            }
            else if (Item is Outlook.SharingItem)
            {
                Outlook.SharingItem sharing = Item as Outlook.SharingItem;
                return GetSenderInformation(sharing);
            }
            else if (Item is Outlook.TaskItem)
            {
                Outlook.TaskItem task = Item as Outlook.TaskItem;
                return GetSenderInformation(task);
            }
            else if (Item is Outlook.TaskRequestItem)
            {
                Outlook.TaskRequestItem taskRequest = Item as Outlook.TaskRequestItem;
                return GetSenderInformation(taskRequest);
            }
            else if (Item is Outlook.TaskRequestAcceptItem)
            {
                Outlook.TaskRequestAcceptItem taskRequestAcceptItem = Item as Outlook.TaskRequestAcceptItem;
                string mailHeader = GetMailHeader(taskRequestAcceptItem.PropertyAccessor);
                return GetSenderInformationFromMailHeader(mailHeader);
            }
            else if (Item is Outlook.TaskRequestDeclineItem)
            {
                Outlook.TaskRequestDeclineItem taskRequestDeclineItem = Item as Outlook.TaskRequestDeclineItem;
                string mailHeader = GetMailHeader(taskRequestDeclineItem.PropertyAccessor);
                return GetSenderInformationFromMailHeader(mailHeader);
            }
            else
            {
                throw new NotSupportedException("未対応のOutook機能です。");
            }
        }

        #region 各Itemの送信者情報取得
        /// <summary>
        /// MailItemの送信者情報(Dto)を取得する
        /// </summary>
        /// <param name="Item">MailItemオブジェクト</param>
        /// <returns>送信者の宛先情報DTO（送信者が取得できない場合null）</returns>
        private static RecipientInformationDto GetSenderInformation(Outlook.MailItem item)
        {
            Outlook.AddressEntry addressEntry;
            Outlook.ExchangeUser exchUser;
            Outlook.Recipient recipient;
            Outlook.ContactItem contactItem;

            //Office365(Exchangeユーザー)からのメール
            if (item.Sender != null)
            {
                addressEntry = item.Sender;
                recipient = Globals.ThisAddIn.Application.Session.CreateRecipient(addressEntry.Address);
                exchUser = getExchangeUser(recipient.AddressEntry);
            }
            //送信元メールアドレスが取れた場合
            else if (item.SenderEmailAddress != null)
            {
                recipient = Globals.ThisAddIn.Application.Session.CreateRecipient(item.SenderEmailAddress);
                addressEntry = recipient.AddressEntry;
                exchUser = getExchangeUser(recipient.AddressEntry);
            }
            //新規にメール作成中の場合ここ
            else
            {
                // 起動されたOutlookのユーザを送信者として取得
                recipient = Globals.ThisAddIn.Application.Session.CurrentUser;
                addressEntry = recipient.AddressEntry;
                exchUser = getExchangeUser(addressEntry);
            }

            //個人の「連絡先」に登録されているかもしれない
            contactItem = null;
            if (addressEntry != null)
            {
                contactItem = addressEntry.GetContact();
            }

            RecipientInformationDto senderInformation = null;

            // 送信者のExchangeUserが取得できた場合
            if (exchUser != null)
            {
                senderInformation = new RecipientInformationDto(exchUser.Name,
                                                                exchUser.Department,
                                                                exchUser.CompanyName,
                                                                FormatJobTitle(exchUser.JobTitle),
                                                                Outlook.OlMailRecipientType.olOriginator);
            }
            // Exchangeアドレス帳にないが、「連絡先」にいる場合
            else if (contactItem != null)
            {
                senderInformation = new RecipientInformationDto(contactItem.FullName,
                                                                contactItem.Department,
                                                                contactItem.CompanyName,
                                                                FormatJobTitle(contactItem.JobTitle),
                                                                Outlook.OlMailRecipientType.olOriginator);
            }
            // Exchangeアドレス帳にも「連絡先」にもない場合
            else
            {
                string displayName;
                if (item.SenderName != null && !Utility.IsEmailAddress(item.SenderName))
                    displayName = FormatDisplayNameAndAddress(item.SenderName, item.SenderEmailAddress);
                else if (recipient != null && !Utility.IsEmailAddress(recipient.Name))
                    displayName = GetDisplayNameAndAddress(recipient);
                else if (addressEntry != null)
                    displayName = FormatDisplayNameAndAddress(addressEntry.Name, addressEntry.Address);
                else
                    displayName = FormatDisplayNameAndAddress(item.SenderName, item.SenderEmailAddress);

                senderInformation = new RecipientInformationDto(displayName, Outlook.OlMailRecipientType.olOriginator);
            }

            return senderInformation;
        }

        /// <summary>
        /// MeetingItemの送信者情報(Dto)を取得する
        /// </summary>
        /// <param name="Item">MeetingItemオブジェクト</param>
        /// <returns>送信者の宛先情報DTO（送信者が取得できない場合null）</returns>
        private static RecipientInformationDto GetSenderInformation(Outlook.MeetingItem item)
        {
            // SenderEmailAddressから、送信者のAddressEntry及びExchangeUserを取得
            Outlook.Recipient recipient = Globals.ThisAddIn.Application.Session.CreateRecipient(item.SenderEmailAddress);

            Outlook.AddressEntry addressEntry = null;
            Outlook.ExchangeUser exchUser = null;
            Outlook.ContactItem contactItem = null;
            if (recipient != null)
            {
                addressEntry = recipient.AddressEntry;
                exchUser = getExchangeUser(addressEntry);
                contactItem = addressEntry.GetContact();
            }

            RecipientInformationDto senderInformation = null;

            // 送信者のExchangeUserが取得できた場合
            if (exchUser != null)
            {
                senderInformation = new RecipientInformationDto(exchUser.Name,
                                                                exchUser.Department,
                                                                exchUser.CompanyName,
                                                                FormatJobTitle(exchUser.JobTitle),
                                                                Outlook.OlMailRecipientType.olOriginator);
            }
            // Exchangeアドレス帳にないが、「連絡先」にいる場合
            else if (contactItem != null)
            {
                senderInformation = new RecipientInformationDto(contactItem.FullName,
                                                                contactItem.Department,
                                                                contactItem.CompanyName,
                                                                FormatJobTitle(contactItem.JobTitle),
                                                                Outlook.OlMailRecipientType.olOriginator);
            }
            // Exchangeアドレス帳にも「連絡先」にもない場合
            else
            {
                string displayName;
                if (item.SenderName != null && !Utility.IsEmailAddress(item.SenderName))
                    displayName = FormatDisplayNameAndAddress(item.SenderName, item.SenderEmailAddress);
                else if (recipient != null && !Utility.IsEmailAddress(recipient.Name))
                    displayName = GetDisplayNameAndAddress(recipient);
                else if (addressEntry != null)
                    displayName = FormatDisplayNameAndAddress(addressEntry.Name, addressEntry.Address);
                else
                    displayName = FormatDisplayNameAndAddress(item.SenderName, item.SenderEmailAddress);

                senderInformation = new RecipientInformationDto(displayName, Outlook.OlMailRecipientType.olOriginator);
            }
            return senderInformation;
        }

        /// <summary>
        /// AppointmentItemの送信者情報(Dto)を取得する
        /// </summary>
        /// <param name="Item">AppointmentItemオブジェクト</param>
        /// <returns>送信者の宛先情報DTO（送信者が取得できない場合null）</returns>
        private static RecipientInformationDto GetSenderInformation(Outlook.AppointmentItem item)
        {
            // 先頭(Recipients[1])のRecipientは送信者なので、送信者のExchangeUserを取得
            Outlook.Recipient recipient = item.Recipients[1];
            Outlook.AddressEntry addressEntry = recipient.AddressEntry;
            Outlook.ExchangeUser exchUser = getExchangeUser(addressEntry);
            if (exchUser == null)
            {
                // 起動されたOutlookのユーザを送信者として取得
                addressEntry = Globals.ThisAddIn.Application.Session.CurrentUser.AddressEntry;
                exchUser = getExchangeUser(addressEntry);
            }

            RecipientInformationDto senderInformation = null;

            // 送信者のExchangeUserが取得できた場合
            if (exchUser != null)
            {
                senderInformation = new RecipientInformationDto(exchUser.Name,
                                                                exchUser.Department,
                                                                exchUser.CompanyName,
                                                                FormatJobTitle(exchUser.JobTitle),
                                                                Outlook.OlMailRecipientType.olOriginator);
            }
            else
            {
                string displayName;
                if (recipient != null && !Utility.IsEmailAddress(recipient.Name))
                    displayName = GetDisplayNameAndAddress(recipient);
                else if (addressEntry != null)
                    displayName = FormatDisplayNameAndAddress(addressEntry.Name, addressEntry.Address);
                else
                    displayName = FormatDisplayNameAndAddress(recipient.Name, recipient.Address);

                senderInformation = new RecipientInformationDto(displayName, Outlook.OlMailRecipientType.olOriginator);
            }

            return senderInformation;
        }

        /// <summary>
        /// ReportItemの送信者情報(Dto)を取得する
        /// </summary>
        /// <param name="Item">ReportItemオブジェクト</param>
        /// <returns>送信者の宛先情報DTO（送信者が取得できない場合null）</returns>
        private static RecipientInformationDto GetSenderInformation(Outlook.ReportItem item)
        {
            RecipientInformationDto senderInformation = new RecipientInformationDto("Microsoft Outlook", Outlook.OlMailRecipientType.olOriginator);
            return senderInformation;
        }

        /// <summary>
        /// SharingItemの送信者情報(Dto)を取得する
        /// </summary>
        /// <param name="Item">SharingItemオブジェクト</param>
        /// <returns>送信者の宛先情報DTO（送信者が取得できない場合null）</returns>
        private static RecipientInformationDto GetSenderInformation(Outlook.SharingItem item)
        {
            if (item.SenderEmailAddress == null)
                return GetCurrentUserInformation();

            Outlook.Recipient recipient = Globals.ThisAddIn.Application.Session.CreateRecipient(item.SenderEmailAddress);

            Outlook.AddressEntry addressEntry = null;
            Outlook.ExchangeUser exchUser = null;
            if (recipient != null)
            {
                addressEntry = recipient.AddressEntry;
                exchUser = getExchangeUser(addressEntry);
            }

            RecipientInformationDto senderInformation = null;

            // 送信者のExchangeUserが取得できた場合
            if (exchUser != null)
            {
                senderInformation = new RecipientInformationDto(exchUser.Name,
                                                                exchUser.Department,
                                                                exchUser.CompanyName,
                                                                FormatJobTitle(exchUser.JobTitle),
                                                                Outlook.OlMailRecipientType.olOriginator);
            }
            // ExchangeUserが取得できないが、送信者情報は取得できた場合
            else
            {
                string displayName;
                if (item.SenderName != null && !Utility.IsEmailAddress(item.SenderName))
                    displayName = FormatDisplayNameAndAddress(item.SenderName, item.SenderEmailAddress);
                else if (recipient != null && !Utility.IsEmailAddress(recipient.Name))
                    displayName = GetDisplayNameAndAddress(recipient);
                else if (addressEntry != null)
                    displayName = FormatDisplayNameAndAddress(addressEntry.Name, addressEntry.Address);
                else
                    displayName = FormatDisplayNameAndAddress(item.SenderName, item.SenderEmailAddress);

                senderInformation = new RecipientInformationDto(displayName, Outlook.OlMailRecipientType.olOriginator);
            }
            return senderInformation;
        }

        /// <summary>
        /// TaskItemの送信者情報(Dto)を取得する
        /// </summary>
        /// <param name="Item">TaskItemオブジェクト</param>
        /// <returns>送信者の宛先情報DTO（送信者が取得できない場合null）</returns>
        private static RecipientInformationDto GetSenderInformation(Outlook.TaskItem item)
        {
            Outlook.Recipient recipient;
            if (Utility.IsSendTaskRequest(item))
            {
                recipient = Globals.ThisAddIn.Application.Session.CurrentUser;
            }
            else
            {
                if (item.Delegator != null)
                    recipient = Globals.ThisAddIn.Application.Session.CreateRecipient(item.Delegator);
                else
                    recipient = null;
            }
            Outlook.ExchangeUser exchUser = null;
            Outlook.AddressEntry addressEntry = null;
            Outlook.ContactItem contactItem = null;
            if (recipient != null)
            {
                addressEntry = recipient.AddressEntry;
                exchUser = getExchangeUser(addressEntry);
            }

            RecipientInformationDto senderInformation = null;

            // 送信者のExchangeUserが取得できた場合
            if (exchUser != null)
            {
                senderInformation = new RecipientInformationDto(exchUser.Name,
                                                                exchUser.Department,
                                                                exchUser.CompanyName,
                                                                FormatJobTitle(exchUser.JobTitle),
                                                                Outlook.OlMailRecipientType.olOriginator);
            }
            // Exchangeアドレス帳にないが、「連絡先」にいる場合
            else if (contactItem != null)
            {
                senderInformation = new RecipientInformationDto(contactItem.FullName,
                                                                contactItem.Department,
                                                                contactItem.CompanyName,
                                                                FormatJobTitle(contactItem.JobTitle),
                                                                Outlook.OlMailRecipientType.olOriginator);
            }
            // Exchangeアドレス帳にも「連絡先」にもない場合
            else
            {
                string displayName;
                if (recipient != null && !Utility.IsEmailAddress(recipient.Name))
                    displayName = GetDisplayNameAndAddress(recipient);
                else if (addressEntry != null)
                    displayName = FormatDisplayNameAndAddress(addressEntry.Name, addressEntry.Address);
                else if (recipient != null)
                    displayName = recipient.Name;
                else
                    displayName = "※取得できませんでした"; //TODO:メールヘッダーのFromから取得する

                senderInformation = new RecipientInformationDto(displayName, Outlook.OlMailRecipientType.olOriginator);
            }
            return senderInformation;
        }

        /// <summary>
        /// TaskRequestItemの送信者情報(Dto)を取得する
        /// </summary>
        /// <param name="Item">TaskRequestItemオブジェクト</param>
        /// <returns>送信者の宛先情報DTO（送信者が取得できない場合null）</returns>
        private static RecipientInformationDto GetSenderInformation(Outlook.TaskRequestItem item)
        {
            Outlook.Recipient recipient;
            Outlook.ContactItem contactItem;
            if (IsSendTaskRequest(item.GetAssociatedTask(false)))
            {
                //これから送信するTaskRequestItem
                recipient = Globals.ThisAddIn.Application.Session.CurrentUser;
                contactItem = null;
            }
            else
            {
                //受信したTaskRequestItem
                Outlook.TaskItem task = item.GetAssociatedTask(false);
                recipient = task.Recipients[1];
                contactItem = recipient.AddressEntry.GetContact();
            }
            Outlook.AddressEntry addressEntry = recipient.AddressEntry;
            Outlook.ExchangeUser exchUser = getExchangeUser(addressEntry);

            RecipientInformationDto senderInformation = null;

            // 送信者のExchangeUserが取得できた場合
            if (exchUser != null)
            {
                senderInformation = new RecipientInformationDto(exchUser.Name,
                                                                exchUser.Department,
                                                                exchUser.CompanyName,
                                                                FormatJobTitle(exchUser.JobTitle),
                                                                Outlook.OlMailRecipientType.olOriginator);
            }
            // Exchangeアドレス帳にないが、「連絡先」にいる場合
            else if (contactItem != null)
            {
                senderInformation = new RecipientInformationDto(contactItem.FullName,
                                                                contactItem.Department,
                                                                contactItem.CompanyName,
                                                                FormatJobTitle(contactItem.JobTitle),
                                                                Outlook.OlMailRecipientType.olOriginator);
            }
            // Exchangeアドレス帳にも「連絡先」にもない場合
            else
            {
                string displayName;
                if (recipient != null && !Utility.IsEmailAddress(recipient.Name))
                    displayName = GetDisplayNameAndAddress(recipient);
                else if (addressEntry != null)
                    displayName = FormatDisplayNameAndAddress(addressEntry.Name, addressEntry.Address);
                else if (recipient != null)
                    displayName = recipient.Name;
                else
                    displayName = "※取得できませんでした"; //TODO:メールヘッダーのFromから取得する

                senderInformation = new RecipientInformationDto(displayName, Outlook.OlMailRecipientType.olOriginator);
            }
            return senderInformation;
        }

        /// <summary>
        /// 自分自身の送信者情報(Dto)を取得する
        /// </summary>
        /// <returns>送信者の宛先情報DTO（送信者が取得できない場合null）</returns>
        private static RecipientInformationDto GetCurrentUserInformation()
        {
            Outlook.Recipient recipient = Globals.ThisAddIn.Application.Session.CurrentUser;
            Outlook.AddressEntry addressEntry = recipient.AddressEntry;
            Outlook.ExchangeUser exchUser = getExchangeUser(addressEntry);

            RecipientInformationDto senderInformation = null;

            // 送信者のExchangeUserが取得できた場合
            if (exchUser != null)
            {
                senderInformation = new RecipientInformationDto(exchUser.Name,
                                                                exchUser.Department,
                                                                exchUser.CompanyName,
                                                                FormatJobTitle(exchUser.JobTitle),
                                                                Outlook.OlMailRecipientType.olOriginator);
            }
            // Exchangeアドレス帳にない場合
            else
            {
                string displayName;
                if (recipient != null && !Utility.IsEmailAddress(recipient.Name))
                    displayName = GetDisplayNameAndAddress(recipient);
                else if (addressEntry != null)
                    displayName = FormatDisplayNameAndAddress(addressEntry.Name, addressEntry.Address);
                else if (recipient != null)
                    displayName = recipient.Name;
                else
                    displayName = "※取得できませんでした";

                senderInformation = new RecipientInformationDto(displayName, Outlook.OlMailRecipientType.olOriginator);
            }
            return senderInformation;
        }
        #endregion

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

        /// <summary>
        /// 表示用に"名前<メールアドレス>"の形式の文字列を取得する
        /// </summary>
        /// <param name="recipient">Recipientオブジェクト</param>
        /// <returns>表示名</returns>
        public static string GetDisplayNameAndAddress(Outlook.Recipient recipient)
        {
            if (recipient.Name == null)
            {
                return recipient.Address;
            }

            //表示名がメールアドレスになっていたら、メールアドレスだけを表示する
            if (recipient.Name != null && recipient.Address != null)
            {
                if (recipient.Name.CompareTo(recipient.Address) == 0)
                {
                    return recipient.Address;
                }
            }


            //表示名内のカッコの中をチェックする
            //カッコの中がメールアドレスだったらメールアドレスが重複表示しないように、表示名をそのまま返す
            System.Text.RegularExpressions.Regex pattern = new System.Text.RegularExpressions.Regex(@".+\((.+)\)");
            System.Text.RegularExpressions.Match match = pattern.Match(recipient.Name);
            if (match.Success)
            {
                string detail = match.Groups[1].Value;
                if (IsEmailAddress(detail))
                {
                    return recipient.Name;
                }
            }
            
            //表示名にカッコがなかったり、カッコの中がメールアドレスではなければ、
            //"表示名 <メールアドレス>"の形式を返す
            string displayName = FormatDisplayNameAndAddress(recipient.Name, recipient.Address);
            return displayName;
        }

        /// <summary>
        /// 文字列がEメールアドレスかどうか判定する
        /// </summary>
        /// <param name="address">判定したい文字列</param>
        /// <returns>true:Eメールアドレス false:Eメールアドレスではない</returns>
        public static bool IsEmailAddress(string address)
        {
            try
            {
                System.Net.Mail.MailAddress mailAddress = new System.Net.Mail.MailAddress(address);
            }
            catch (FormatException)
            {
                return false;
            }
            return true;
        }

        /// <summary>
        /// 名前とアドレスを表示用に整形する
        /// </summary>
        /// <param name="Name">名前</param>
        /// <param name="MailAddress">メールアドレス</param>
        /// <returns>表示名</returns>
        public static string FormatDisplayNameAndAddress(string Name, string MailAddress)
        {
            if (Name != null && MailAddress != null)
            {
                //表示名がメールアドレスになっていたら、メールアドレスだけを表示する
                if (MailAddress.CompareTo(Name) == 0)
                    return MailAddress;
                else
                    return string.Format("{0} <{1}>", Name, MailAddress);
            }
            else if (MailAddress == null)
                return Name;
            else if (Name == null)
                return MailAddress;
            else
                return null;
        }

        /// <summary>
        /// AddressEntryを元にxchangeUserオブジェクトを取得する
        /// </summary>
        /// <param name="entry">AddressEntryオブジェクト</param>
        /// <returns>AddressEntryに紐づいたExchangeUserオブジェクト。失敗した場合はnullを返す。</returns>
        private static Outlook.ExchangeUser getExchangeUser(Outlook.AddressEntry entry)
        {
            Outlook.ExchangeUser exchUser;
            try
            {
                exchUser = entry.GetExchangeUser();
            }
            catch (Exception)
            {
                exchUser = null;
            }
            return exchUser;
        }

        public static bool IsSendTaskRequest(Outlook.TaskItem item)
        {
            if (item.Delegator == null && item.Owner == null && item.Saved)
            {
                return false;
            }
            else if (item.Delegator != null && item.Delegator.CompareTo(item.Owner) != 0)
            {
                //受信したTaskRequestItem
                return false;
            }
            else
            {
                //これから送信するTaskRequestItem
                return true;
            }
        }

        public static string GetMailHeader(Outlook.PropertyAccessor prop)
        {
            const string PR_TRANSPORT_MESSAGE_HEADERS = "http://schemas.microsoft.com/mapi/proptag/0x007D001E";
            string headerString = (string)prop.GetProperty(PR_TRANSPORT_MESSAGE_HEADERS);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(prop);

            return headerString;
        }

        public static RecipientInformationDto GetSenderInformationFromMailHeader(string MailHeader)
        {
            string[] headerStrings = MailHeader.Replace("\r\n", "\n").Split('\n');

            for (int i=0; i<headerStrings.Length; i++)
            {
                string[] headerItems = headerStrings[i].Split(':');
                if (headerItems[0].CompareTo("From") != 0)
                {
                    continue;
                }
                //『From』には以下のように入っている
                //From: =?utf-8?B?TmFrYW5pc2hpIFl1bmEgKOS4reilvyDmgqDoj5wp?=
                //\t < yuna.nakanishi@jp.ricoh.com >

                //表示名を取得
                string[] froms = headerItems[1].Split('?');
                byte[] fromBytes = Convert.FromBase64String(froms[3]);
                string fromName = Encoding.UTF8.GetString(fromBytes);

                Outlook.Recipient recipient = Globals.ThisAddIn.Application.Session.CreateRecipient(fromName);

                string fromAddress = null;
                if (recipient == null)
                {
                    //メールアドレスを取得
                    System.Text.RegularExpressions.Regex pattern = new System.Text.RegularExpressions.Regex(@".+\<(.+)\>");
                    System.Text.RegularExpressions.Match match = pattern.Match(headerStrings[i + 1]);
                    if (match.Success)
                    {
                        string detail = match.Groups[1].Value;
                        if (IsEmailAddress(detail))
                        {
                            fromAddress = detail;
                            recipient = Globals.ThisAddIn.Application.Session.CreateRecipient(fromAddress);
                        }
                    }
                }

                if (recipient != null)
                {
                    Outlook.ExchangeUser exchUser = recipient.AddressEntry.GetExchangeUser();
                    if (exchUser != null)
                    {
                        return new RecipientInformationDto(exchUser.Name,
                                                           exchUser.Department,
                                                           exchUser.CompanyName,
                                                           FormatJobTitle(exchUser.JobTitle),
                                                           Outlook.OlMailRecipientType.olOriginator);

                    }

                    Outlook.ContactItem contactItem = recipient.AddressEntry.GetContact();
                    if (contactItem != null)
                    {
                        new RecipientInformationDto(contactItem.FullName,
                                                    contactItem.Department,
                                                    contactItem.CompanyName,
                                                    FormatJobTitle(contactItem.JobTitle),
                                                    Outlook.OlMailRecipientType.olOriginator);
                    }

                    string displayName;
                    if (!Utility.IsEmailAddress(recipient.Name))
                        displayName = GetDisplayNameAndAddress(recipient);
                    else if (recipient.AddressEntry != null)
                        displayName = FormatDisplayNameAndAddress(recipient.AddressEntry.Name, recipient.AddressEntry.Address);
                    else if (recipient != null)
                        displayName = recipient.Name;
                    else
                        displayName = "※取得できませんでした";

                    return new RecipientInformationDto(displayName, Outlook.OlMailRecipientType.olOriginator);

                }
                else
                {
                    string displayName;
                    if (fromName != null && fromAddress != null)
                        displayName = FormatDisplayNameAndAddress(fromName, fromAddress);
                    else if (fromName != null)
                        displayName = fromName;
                    else if (fromAddress != null)
                        displayName = fromAddress;
                    else
                        displayName = "※取得できませんでした";

                    return new RecipientInformationDto(displayName, Outlook.OlMailRecipientType.olOriginator);
                }
            }
            return new RecipientInformationDto("※取得できませんでした", Outlook.OlMailRecipientType.olOriginator);
        }
    }
}
