using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OutlookRecipientConfirmationAddin
{
    /// <summary>
    /// 色々な連絡先(O365, Notes...)のインターフェース
    /// </summary>
    interface IContact
    {

        /// <summary>
        /// 連絡先を取得する抽象メソッド
        /// </summary>
        /// <returns></returns>
        ContactItem getContactItem();
            
      
    }
}
