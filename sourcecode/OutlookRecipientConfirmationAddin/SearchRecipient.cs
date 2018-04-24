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
        /// <summary>
        /// メールのアドレスから宛先情報を検索する
        /// </summary>
        /// <param name="addressList">メールのTO, CC, BCC</param> 
        /// <returns> 検索した宛先情報のリスト</returns>
        public List<RecipientInformationDto> SearchContact(List<Recipient> recipientsList, Utility.OutlookItemType itemType)
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

                try
                {
                    /// それぞれの連絡先クラスで検索する
                    foreach (var item in contactList)
                    {
                        ContactItem contactItem = item.getContactItem(recipient);

                        /// 送信先アドレスからその人の情報が見つかれば、名、部署、会社名、タイプをDtoにセット
                        if (contactItem != null)
                        {
                            string jobTitle = Utility.FormatJobTitle(contactItem.JobTitle);

                            // TaskRequstResponseは強制的にToにする
                            OlMailRecipientType recipientType = itemType != Utility.OutlookItemType.TaskRequestResponse ?
                                                                    (OlMailRecipientType)recipient.Type :
                                                                    OlMailRecipientType.olTo;

                            recipientInformation = new RecipientInformationDto(contactItem.FullName, contactItem.Department, contactItem.CompanyName, jobTitle, recipientType);
                            break;
                        }
                    }

                    if (recipientInformation == null)
                    {
                        recipientInformation = new RecipientInformationDto(
                            Utility.GetDisplayNameAndAddress(recipient),
                            (OlMailRecipientType)recipient.Type);
                    }
                    _recipientInformationList.Add(recipientInformation);

                }
                /// 例外が発生した場合、Nameを表示する
                catch (System.Exception ex)
                {
                    Console.WriteLine(ex.Message);

                    _recipientInformationList.Add(new RecipientInformationDto(
                        Utility.GetDisplayNameAndAddress(recipient),
                        (OlMailRecipientType)recipient.Type));
                }
            }

            return _recipientInformationList;
        }


    }
}
