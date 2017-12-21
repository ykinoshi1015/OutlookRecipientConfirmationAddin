using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OutlookRecipientConfirmationAddin
{
    /// <summary>
    /// 具象クラスの連絡先インスタンスを生成するFactoryクラス
    /// </summary>
    class ContactFactory
    {

        ///  全ての連絡先クラスのリスト
        private List<IContact> contactClassList = new List<IContact>();

        public List<IContact> CreateContacts()
        {
            IContact office365Contact = new Office365Contact();
            contactClassList.Add(office365Contact);

            /// O365以外のクラスがあればここに追加

            return contactClassList;
        }

    }
}
