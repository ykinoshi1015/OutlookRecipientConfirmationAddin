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
        public ContactItem getContactItem()
        {
            ///CreateItem
            
            /// 連絡先を持ってくる処理
            return contact;
        }
    }
}
