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
        List<Recipient> recipientsList;

        /// 検索結果の宛先情報のリスト
        private List<RecipientInformationDto> RecipientInformationList = new List<RecipientInformationDto>();

        /// コンストラクタ
        public SearchRecipient(List<Recipient> recipientsList)
        {
            this.recipientsList = recipientsList;
        }

        /// <summary>
        /// メールのアドレスから宛先情報を検索する
        /// </summary>
        /// <param name="addressList"></param> メールのTO, CC, BCC
        /// <returns> 検索した宛先情報のリスト</returns>
        public List<RecipientInformationDto> SearchContact()
        {
            /// ファクトリオブジェクトに連絡先クラスのインスタンスの生成をしてもらう
            ContactFactory contactFactory = new ContactFactory();
            List<IContact> contactList = contactFactory.CreateContacts();

            /// ある1人の受信者の宛先情報を取得する
            foreach (var recipient in recipientsList)
            {

                /// それぞれの連絡先クラスで検索する
                foreach (var item in contactList)
                {
                    ContactItem contactItem = item.getContactItem(recipient);

                    /// 送信先アドレスから、その人の情報が見つかればDtoにセット
                    if (contactItem != null)
                    {

                        String fullName = contactItem.FullName;
                        String division = contactItem.Department;
                        String companyName = contactItem.CompanyName;
                        OlMailRecipientType recipientType = (OlMailRecipientType)recipient.Type;

                        RecipientInformationDto recipientInformation = new RecipientInformationDto(fullName, division, companyName, recipientType);
                        RecipientInformationList.Add(recipientInformation);
                    }

                    /// このアドレスの検索が完了
                    break;
                }
               
            }
            
            return RecipientInformationList;
        }


    }
}
