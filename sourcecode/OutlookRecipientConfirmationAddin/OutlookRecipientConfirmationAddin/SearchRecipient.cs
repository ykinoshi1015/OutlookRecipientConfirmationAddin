using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OutlookRecipientConfirmationAddin
{
    /// <summary>
    /// 送信されるメールから取得したTO, CC, BCCをアドレスリストから検索するクラス
    /// </summary>
    class SearchRecipient
    {
        /// TO, CC, BCCに入力されたアドレスのリスト
        List<String> toList;
        List<String> ccList;
        List<String> bccList;

        /// 検索結果の宛先情報のリスト
        private List<RecipientInformationDto> RecipientInformationList = new List<RecipientInformationDto>();

        /// コンストラクタ
        public SearchRecipient(List<String> toList, List<String> ccList, List<String> bccList)
        {
            this.toList = toList;
            this.ccList = ccList;
            this.bccList = bccList;
        }

        /// <summary>
        /// メールのアドレスから宛先情報を検索する
        /// </summary>
        /// <param name="addressList"></param> メールのTO, CC, BCC
        /// <returns> 検索した宛先情報のリスト</returns>
        public List<RecipientInformationDto> SearchContact(List<String> addressList)
        {
            /// ファクトリオブジェクトに連絡先クラスのインスタンスの生成をしてもらう
            ContactFactory contactFactory = new ContactFactory();
            List<IContact> contactList = contactFactory.CreateContacts();

            /// toの宛先情報を取得する
            foreach (var address in toList)
            {

                /// それぞれの連絡先クラスで検索する
                foreach (var item in contactList)
                {
                    List<ContactItem> contactItemList = item.getContactItem();

                    /// 連絡先クラスにあるすべての連絡先から検索
                    foreach (var contact in contactItemList)
                    {
                        /// addressととってきたcontactの連絡先が一致したら、RecipientInformationDtoにセット
                        if (contact.Email1Address.Equals(address))
                        {
                            String fullName = contact.LastNameAndFirstName;
                            String division = contact.Department;
                            String companyName = contact.CompanyName;
                            OlMailRecipientType recipientType = OlMailRecipientType.olTo;

                            RecipientInformationDto recipientInformation = new RecipientInformationDto(fullName, division, companyName, recipientType);
                            RecipientInformationList.Add(recipientInformation);

                            /// このアドレスの検索が完了
                            goto ExitLoop;
                        }
                    }
                }

                ExitLoop:;

            }
            return RecipientInformationList;
        }
    }
}
