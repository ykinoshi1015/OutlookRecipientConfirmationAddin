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

            String fullName;
            String division;
            String companyName;
            OlMailRecipientType recipientType;
            String jobTitle;
            String emailAddress;

            /// ある1人の受信者の宛先情報を取得する
            foreach (var recipient in recipientsList)
            {

                /// それぞれの連絡先クラスで検索する
                foreach (var item in contactList)
                {
                    ContactItem contactItem = item.getContactItem(recipient);

                    /// 送信先アドレスからその人の情報が見つかれば、名、部署、会社名、タイプをDtoにセット
                    if (contactItem.FullName != null)
                    {
                        fullName = contactItem.FullName;
                        division = contactItem.Department;
                        companyName = contactItem.CompanyName;
                        recipientType = (OlMailRecipientType)recipient.Type;
                        emailAddress = "";

                        /// 表示する役職ならDtoに入れる、違えば空文字を入れる
                        if (contactItem.JobTitle != "" && contactItem.JobTitle != "担当")
                        {
                            jobTitle = contactItem.JobTitle;
                        }
                        else
                        {
                            jobTitle = null;
                        }
                        
                    }
                    /// アドレス帳に登録されていない時は、タイプとメールアドレスをDtoにセット
                    else
                    {
                        fullName = null;
                        division = null;
                        companyName = null;
                        jobTitle = null;
                        recipientType = (OlMailRecipientType)recipient.Type;
                        emailAddress = recipient.Address;
                    }

                    RecipientInformationDto recipientInformation = new RecipientInformationDto(fullName, division, companyName, recipientType, jobTitle, emailAddress);
                    RecipientInformationList.Add(recipientInformation);

                    /// このアドレスの検索が完了
                    break;
                }

            }

            return RecipientInformationList;
        }


    }
}
