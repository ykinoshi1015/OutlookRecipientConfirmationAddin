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

        /// 検索結果の宛先情報のリスト
        private List<RecipientInformationDto> RecipientInformationList = new List<RecipientInformationDto>();

        /// <summary>
        /// メールのアドレスから宛先情報を検索する
        /// </summary>
        /// <param name="addressList"></param> メールのTO, CC, BCC
        /// <returns> 検索した宛先情報のリスト</returns>
        public List<RecipientInformationDto> SearchContact(List<Recipient> recipientsList)
        {
            /// ファクトリオブジェクトに連絡先クラスのインスタンスの生成をしてもらう
            ContactFactory contactFactory = new ContactFactory();
            List<IContact> contactList = contactFactory.CreateContacts();

            /// ある1人の受信者の宛先情報を取得する
            foreach (var recipient in recipientsList)
            {
                RecipientInformationDto recipientInformation = null;

                /// それぞれの連絡先クラスで検索する
                foreach (var item in contactList)
                {
                    ContactItem contactItem = item.getContactItem(recipient);

                    /// 送信先アドレスからその人の情報が見つかれば、名、部署、会社名、タイプをDtoにセット
                    if (contactItem != null)
                    {
                        /// 表示する役職ならDtoに入れる、違えば空文字を入れる
                        string jobTitle;
                        if (contactItem.JobTitle != "" && contactItem.JobTitle != "担当")
                        {
                            jobTitle = contactItem.JobTitle;
                        }
                        else
                        {
                            jobTitle = null;
                        }
                        recipientInformation = new RecipientInformationDto(contactItem.FullName, contactItem.Department, contactItem.CompanyName, jobTitle, (OlMailRecipientType)recipient.Type);
                        break;
                    }
                }
                if (recipientInformation == null)
                {
                    recipientInformation = new RecipientInformationDto(recipient.Address, (OlMailRecipientType)recipient.Type);
                }
                RecipientInformationList.Add(recipientInformation);
            }

            return RecipientInformationList;
        }


    }
}
