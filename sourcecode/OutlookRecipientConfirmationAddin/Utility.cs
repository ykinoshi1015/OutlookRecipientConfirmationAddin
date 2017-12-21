using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OutlookRecipientConfirmationAddin
{
    /// <summary>
    /// 宛先情報を表示用にフォーマッティングするクラス
    /// </summary>
    class Utility
    {

        public static string Formatting(RecipientInformationDto recipientInformation)
        {
            string formattedRecipient;

            /// 受信者の情報が見つかったとき
            if (recipientInformation.fullName != null && !recipientInformation.fullName.Equals(""))
            {
                formattedRecipient = string.Format("{0} {1} ({2}【{3}】)", recipientInformation.fullName, recipientInformation.jobTitle, recipientInformation.division, recipientInformation.companyName);
            }
            /// 受信者の情報が見つからなかったとき
            else
            {
                formattedRecipient = recipientInformation.emailAddress;
            }

            return formattedRecipient;
        }
    }
}
