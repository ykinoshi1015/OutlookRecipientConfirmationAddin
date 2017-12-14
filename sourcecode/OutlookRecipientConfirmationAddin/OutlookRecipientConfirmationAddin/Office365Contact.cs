using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;
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
        List<ContactItem> contactList = null;

        /// O365にある連絡先を全部返す
        public List<ContactItem> getContactItem()
        {
            /// 連絡先フォルダ―から情報をもってくる
            Application application = new Application(); ///newしていいの？
            NameSpace outlookNameSpace = application.GetNamespace("MAPI");
            MAPIFolder contactsFolder = outlookNameSpace.GetDefaultFolder(OlDefaultFolders.olFolderContacts);
            Items contactItems = contactsFolder.Items;

            /// リストに入れる
            foreach(var contact in contactItems)
            {
                contactList.Add(contact as ContactItem);
            }

            return contactList;
        }
    }
}
