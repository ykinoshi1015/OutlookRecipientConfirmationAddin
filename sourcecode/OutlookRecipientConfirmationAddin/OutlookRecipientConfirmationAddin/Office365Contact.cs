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
        ContactItem contactItem;

        /// O365にある連絡先を全部返す
        public ContactItem getContactItem(Recipient recipient)
        {
            contactItem = Globals.ThisAddIn.Application.CreateItem(OlItemType.olContactItem) as ContactItem;

            /// 連絡先フォルダ―から情報をもってくる
            AddressLists addrLists;
            addrLists = Globals.ThisAddIn.Application.Session.AddressLists;

            ExchangeUser exchUser = recipient.AddressEntry.GetExchangeUser();

            //Exchangeアドレス帳から選択されたユーザーの場合　
            if (exchUser != null)
            {
                Debug.WriteLine(string.Format("{0}/{1}/{2}/{3}",
                    exchUser.Name,
                    exchUser.CompanyName,
                    exchUser.Department,
                    exchUser.JobTitle));

                contactItem.FullName = exchUser.Name;
                contactItem.CompanyName = exchUser.CompanyName;
                contactItem.Department = exchUser.Department;
            }
            else
            {
                //ローカルのアドレス帳から選択されたユーザーの場合(お気に入りリストなど)
                contactItem = SearchInAllUsers(recipient.Address);
            }

            return contactItem;
        }

        private ContactItem SearchInAllUsers(string address)
        {
            Recipient recResolve = Globals.ThisAddIn.Application.Session.CreateRecipient(address);
            ExchangeUser exchUser = recResolve.AddressEntry.GetExchangeUser();
            if (exchUser != null)
            {
                Debug.WriteLine(string.Format("{0} {1} {2} {3}",
                    exchUser.Name,
                    exchUser.CompanyName,
                    exchUser.Department,
                    exchUser.BusinessTelephoneNumber));

                contactItem.FullName = exchUser.Name;
                contactItem.CompanyName = exchUser.CompanyName;
                contactItem.Department = exchUser.Department;
            }
            
            return contactItem;

        }
    }
}
