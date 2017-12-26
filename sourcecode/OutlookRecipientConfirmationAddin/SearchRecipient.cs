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
        /// 職種が担当の場合の定数
        private const string TANTOU = "担当";

        /// <summary>
        /// メールのアドレスから宛先情報を検索する
        /// </summary>
        /// <param name="addressList">メールのTO, CC, BCC</param> 
        /// <returns> 検索した宛先情報のリスト</returns>
        public List<RecipientInformationDto> SearchContact(List<Recipient> recipientsList)
        {

            /// 検索結果の宛先情報のリスト
            List<RecipientInformationDto> _recipientInformationList = new List<RecipientInformationDto>();

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
                    try
                    {
                        ContactItem contactItem = item.getContactItem(recipient);


                        /// 送信先アドレスからその人の情報が見つかれば、名、部署、会社名、タイプをDtoにセット
                        if (contactItem != null)
                        {
                            /// 表示する役職ならDtoに入れる、違えば空文字を入れる
                            string jobTitle = contactItem.JobTitle;
                            if (TANTOU.Equals(contactItem.JobTitle) || contactItem.JobTitle == null)
                            {
                                jobTitle = "";
                            }
                            recipientInformation = new RecipientInformationDto(contactItem.FullName, contactItem.Department, contactItem.CompanyName, jobTitle, (OlMailRecipientType)recipient.Type);
                            break;
                        }

                    }
                    /// 例外が発生した場合、Nameを表示する
                    catch (System.Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                        recipientInformation = new RecipientInformationDto((OlMailRecipientType)recipient.Type, recipient.Name);
                    }
                }

                if (recipientInformation == null)
                {
                    recipientInformation = new RecipientInformationDto(recipient.Address, (OlMailRecipientType)recipient.Type);
                }
                _recipientInformationList.Add(recipientInformation);
            }

            return _recipientInformationList;
        }


    }
}
