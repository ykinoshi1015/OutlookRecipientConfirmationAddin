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
        public String fullName { get; set; }
        /// 部署
        public String division { get; set; }
        /// 会社名 
        public String companyName { get; set; }
        /// 宛先タイプ
        public OlMailRecipientType recipientType { get; set; }

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
