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

        public String Formatting(RecipientInformationDto recipientInformation)
        {
            var formattedRecipient = new StringBuilder();

            /// 受信者の情報が見つかったとき
            if (recipientInformation.fullName != null) {
                formattedRecipient.Append(recipientInformation.fullName + " ");
                formattedRecipient.Append(recipientInformation.jobTitle);
                formattedRecipient.Append(" (");
                formattedRecipient.Append(recipientInformation.division);
                formattedRecipient.Append("【");
                formattedRecipient.Append(recipientInformation.companyName);
                formattedRecipient.Append("】");
                formattedRecipient.Append(")");
            }
            /// 受信者の情報が見つからなかったとき
            else
            {
                formattedRecipient.Append(recipientInformation.emailAddress);
            }

            System.Diagnostics.Debug.WriteLine(formattedRecipient);
            
            return formattedRecipient.ToString();
        }
    }
}
