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
            
            ///{To:} 
            formattedRecipient.Append(recipientInformation.fullName);
            formattedRecipient.Append("(");
            formattedRecipient.Append(recipientInformation.division);
            formattedRecipient.Append("【");
            formattedRecipient.Append(recipientInformation.companyName);
            formattedRecipient.Append("】");
            formattedRecipient.Append(")");

            System.Diagnostics.Debug.WriteLine(formattedRecipient);


            //            /// 宛名
            //private String fullName { get; set; }
            ///// 部署
            //private String division { get; set; }
            ///// 会社名 
            //private String companyName { get; set; }
            ///// 宛先タイプ
            //private OlMailRecipientType recipientType { get; set; }


            return formattedRecipient.ToString();
        }
    }
}
