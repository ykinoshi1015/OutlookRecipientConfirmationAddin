﻿using Microsoft.Office.Interop.Outlook;
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

            /// 連絡先フォルダ―から情報をもってくる
            AddressLists addrLists;
            addrLists = Globals.ThisAddIn.Application.Session.AddressLists;

            //Exchangeアドレス帳から選択されたユーザーの場合
            ExchangeUser exchUser = recipient.AddressEntry.GetExchangeUser();
            
            if (exchUser == null)
            {
                //ローカルのアドレス帳から選択されたユーザーの場合(お気に入りリストなど)
                Recipient recResolve = Globals.ThisAddIn.Application.Session.CreateRecipient(recipient.Address);
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