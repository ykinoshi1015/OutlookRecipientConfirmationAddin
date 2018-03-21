﻿using System;
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
        public enum OutlookItemType { Mail, Meeting, Appointment, MeetingResponse, Sharing, Report };

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
            bool isAppointmentItem = false;

            Outlook.MailItem mail = Item as Outlook.MailItem;
            // MailItemの場合
            if (mail != null)
            {
                recipients = mail.Recipients;
                type = OutlookItemType.Mail;
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

                }
                // 会議出席依頼を送信する場合など
                else
                {
                    type = OutlookItemType.Meeting;
                }

                recipients = meeting.Recipients;
            }
            // AppointmentItemの場合（編集中/送信されていない状態でトレイにある会議招集メール、開催者が取り消した会議のキャンセル通知（自分承認済み））
            else if (Item is Outlook.AppointmentItem)
            {
                Outlook.AppointmentItem appointment = Item as Outlook.AppointmentItem;

                recipients = appointment.Recipients;
                type = OutlookItemType.Appointment;
                isAppointmentItem = true;
            }
            else if (Item is Outlook.SharingItem)
            {
                Outlook.SharingItem item = Item as Outlook.SharingItem;

                recipients = item.Recipients;
                type = OutlookItemType.Sharing;
            }
            else if (Item is Outlook.ReportItem)
            {
                Outlook.ReportItem item = Item as Outlook.ReportItem;

                //ReportItemのままだと送信先が取れないため、
                //いったんIPM.Noteとして別名保存⇒ロードしてからRecipientsを取得する
                Outlook.ReportItem copiedReport = item.Copy();
                copiedReport.MessageClass = "IPM.Note";
                copiedReport.Save();

                //IPM.Noteとして穂zんしてからロードするとMailItemとして扱えるようになる
                var newReportItem = Globals.ThisAddIn.Application.Session.GetItemFromID(copiedReport.EntryID);
                Outlook.MailItem newMailItem = newReportItem as Outlook.MailItem;
                recipients = newMailItem.Recipients;
                type = OutlookItemType.Report;

                copiedReport.Delete();
            }

            // 受信者の情報をリストに入れる
            List<Outlook.Recipient> recipientsList = new List<Outlook.Recipient>();

            int i = isAppointmentItem ? 2 : 1;

            for (; i <= recipients.Count; i++)
            {
                // recipients[i]がBccまたはリソース
                if (recipients[i].Type == (int)Outlook.OlMailRecipientType.olBCC)
                {
                    // Bccや、選択されたリソースの場合
                    if (recipients[i].Sendable)
                    {
                        recipientsList.Add(recipients[i]);
                    }
                    // 選択されていないリソースの場合
                    else
                    {
                        continue;
                    }
                }
                // 送信者、To、Ccの場合
                else
                {
                    recipientsList.Add(recipients[i]);
                }
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
            Outlook.Recipient recResolve = null;
            Outlook.ContactItem contactItem = null;

            string senderName = null;
            string senderAddress = null;

            // MailItemの場合
            if (Item is Outlook.MailItem)
            {
                Outlook.MailItem mail = (Item as Outlook.MailItem);

                // 送信元のアカウントに対応するSenderプロパティを取得
                sender = mail.Sender;
                if (sender != null)
                {
                    recResolve = Globals.ThisAddIn.Application.Session.CreateRecipient(sender.Address);
                    exchUser = getExchangeUser(recResolve.AddressEntry);
                }
                // 新規メッセージ編集中/送信時はSenderはnullなので、SenderEmailAddressからExchangeUserを探す
                else if (mail.SenderEmailAddress != null)
                {
                    recResolve = Globals.ThisAddIn.Application.Session.CreateRecipient(mail.SenderEmailAddress);
                    exchUser = getExchangeUser(recResolve.AddressEntry);
                }
                if (recResolve != null)
                {
                    contactItem = recResolve.AddressEntry.GetContact();
                }
                senderName = mail.SenderName;
                senderAddress = mail.SenderEmailAddress;
            }
            // MeetingItemの場合
            else if (Item is Outlook.MeetingItem)
            {
                Outlook.MeetingItem meeting = Item as Outlook.MeetingItem;

                // SenderEmailAddressから、送信者のAddressEntry及びExchangeUserを取得
                recResolve = Globals.ThisAddIn.Application.Session.CreateRecipient(meeting.SenderEmailAddress);
                if (recResolve != null)
                {
                    sender = recResolve.AddressEntry;
                    exchUser = getExchangeUser(sender);
                    contactItem = sender.GetContact();
                }
                senderName = meeting.SenderName;
                senderAddress = meeting.SenderEmailAddress;
            }
            // AppointmentItemの場合
            else if (Item is Outlook.AppointmentItem)
            {
                Outlook.AppointmentItem appointment = Item as Outlook.AppointmentItem;

                // 先頭(Recipients[1])のRecipientは送信者なので、送信者のExchangeUserを取得
                sender = appointment.Recipients[1].AddressEntry;
                exchUser = getExchangeUser(sender);
                if (exchUser == null)
                {
                    // 起動されたOutlookのユーザを送信者として取得
                    sender = Globals.ThisAddIn.Application.Session.CurrentUser.AddressEntry;
                    exchUser = getExchangeUser(sender);
                }
                contactItem = sender.GetContact();
            }
            else if (Item is Outlook.ReportItem)
            {
                Outlook.ReportItem report = Item as Outlook.ReportItem;

                senderName = "Microsoft Outlook";
            }
            else if (Item is Outlook.SharingItem)
            {
                Outlook.SharingItem sharing = Item as Outlook.SharingItem;

                // SenderEmailAddressから、送信者のAddressEntry及びExchangeUserを取得
                recResolve = Globals.ThisAddIn.Application.Session.CreateRecipient(sharing.SenderEmailAddress);
                if (recResolve != null)
                {
                    sender = recResolve.AddressEntry;
                    exchUser = getExchangeUser(sender);
                    contactItem = sender.GetContact();
                }
                senderName = sharing.SenderName;
                senderAddress = sharing.SenderEmailAddress;
            }

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
            // ExchangeUserが取得できないが、送信者情報は取得できた場合
            else if (sender != null)
            {
                string displayName;
                if (recResolve != null)
                    displayName = GetDisplayNameAndAddress(recResolve);
                else
                    displayName = FormatDisplayNameAndAddress(sender.Name, sender.Address);

                senderInformation = new RecipientInformationDto(displayName, Outlook.OlMailRecipientType.olOriginator);
            }
            // いずれも失敗した場合
            else if (senderName != null)
            {
                senderInformation = new RecipientInformationDto(
                    FormatDisplayNameAndAddress(senderName, senderAddress),
                    Outlook.OlMailRecipientType.olOriginator);
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

        /// <summary>
        /// 表示用に"名前<メールアドレス>"の形式の文字列を取得する
        /// </summary>
        /// <param name="recipient">Recipientオブジェクト</param>
        /// <returns>表示名</returns>
        public static string GetDisplayNameAndAddress(Outlook.Recipient recipient)
        {
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
                if (isEmailAddress(detail))
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
        public static bool isEmailAddress(string address)
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
            if (MailAddress == null)
                return string.Format("{0}", Name);
            else
                return string.Format("{0} <{1}>", Name, MailAddress);
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
    }
}
