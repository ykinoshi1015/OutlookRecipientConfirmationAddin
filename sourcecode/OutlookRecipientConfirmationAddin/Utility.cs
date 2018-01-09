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

            /// 名前を表示するとき
            if (!recipientInformation.fullName.Equals(""))
            {
                /// Exchangeアドレス帳で受信者の情報が見つかったとき
                if (recipientInformation.division != null)
                {
                    formattedRecipient = string.Format("{0} {1} ({2}【{3}】)", recipientInformation.fullName, recipientInformation.jobTitle, recipientInformation.division, recipientInformation.companyName);
                }
                /// グループ名のみを表示するとき
                else
                {
                    formattedRecipient = recipientInformation.fullName;
                }
            }
            /// 受信者の情報が見つからなかったとき、例外のとき
            else
            {
                /// アドレスだけ表示する
                formattedRecipient = recipientInformation.emailAddress;
            }

            return formattedRecipient;
        }
    }
}
