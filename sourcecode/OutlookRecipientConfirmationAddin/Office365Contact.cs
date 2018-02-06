using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OutlookRecipientConfirmationAddin
{
    /// <summary>
    /// O365の連絡先クラス
    /// </summary>
    class Office365Contact : IContact
    {
        /// O365にある連絡先を全部返す
        public ContactItem getContactItem(Recipient recipient)
        {
            ContactItem contactItem = null;

            /// O365のグループアドレスの場合
            if (OlAddressEntryUserType.olExchangeDistributionListAddressEntry == recipient.AddressEntry.AddressEntryUserType)
            {
                /// グループ名を入れて戻る
                contactItem = Globals.ThisAddIn.Application.CreateItem(OlItemType.olContactItem);
                contactItem.FullName = recipient.Name;
                return contactItem;
            }

            /// Exchangeアドレス帳から選択されたユーザーの場合は、ここで取得
            ExchangeUser exchUser = recipient.AddressEntry.GetExchangeUser();

            /// それ以外の場合
            /// （Outlook連絡先フォルダーのアドレスエントリ、もしくはNotesのクラウドメールアドレスからの受信メールに返信する際のアドレスエントリの場合など）
            if (exchUser == null)
            {
                Recipient recResolve = Globals.ThisAddIn.Application.Session.CreateRecipient(recipient.Address);
                /// Exchangeアドレス帳に存在するアドレスなら、exchUserが見つかる
                exchUser = recResolve.AddressEntry.GetExchangeUser();
            }

            /// ExchangeUserが見つかれば、ContactItemに入れる
            if (exchUser != null)
            {
                contactItem = Globals.ThisAddIn.Application.CreateItem(OlItemType.olContactItem);
                contactItem.FullName = exchUser.Name;
                contactItem.CompanyName = exchUser.CompanyName;
                contactItem.Department = exchUser.Department;
                contactItem.JobTitle = exchUser.JobTitle;
            }

            return contactItem;
        }

    }
}
