using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OutlookRecipientConfirmationAddin
{
    /// <summary>
    /// 宛先情報のDto
    /// </summary>
    class RecipientInformationDto
    {
        /// 宛名
        private String fullName { get; set; }
        /// 部署
        private String division { get; set; }
        /// 会社名 
        private String companyName { get; set; }
        /// 宛先タイプ
        private OlMailRecipientType recipientType { get; set; }

        public RecipientInformationDto(String fullName, String division,
            String companyName, OlMailRecipientType recipientType)
        {
            this.fullName = fullName;
            this.division = division;
            this.companyName = companyName;
            this.recipientType = recipientType;
        }


    }
}
