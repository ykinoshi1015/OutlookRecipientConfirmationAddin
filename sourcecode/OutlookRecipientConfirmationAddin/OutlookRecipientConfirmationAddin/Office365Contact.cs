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
        List<ContactItem> contactList;

        public List<ContactItem> getContactItem()
        {

            /// O365にある連絡先を全部返す
         
            ///CreateItem

            return contactList;
        }
    }
}
