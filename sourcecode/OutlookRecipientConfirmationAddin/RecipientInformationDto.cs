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
    public class RecipientInformationDto
    {
        /// 宛名
        public string fullName { get; set; }
        /// 部署
        public string division { get; set; }
        /// 会社名 
        public string companyName { get; set; }
        /// 宛先タイプ
        public OlMailRecipientType recipientType { get; set; }
        public string jobTitle { get; set; }
        public string emailAddress { get; set; }


        public RecipientInformationDto(string emailAddress, OlMailRecipientType recipientType)
        {
            this.fullName = "";
            this.division = "";
            this.companyName = "";
            this.recipientType = recipientType;
            this.jobTitle = "";
            this.emailAddress = emailAddress;
        }

        public RecipientInformationDto(string fullName, string division,
            string companyName, string jobTitle, OlMailRecipientType recipientType)
        {
            this.fullName = fullName;
            this.division = division;
            this.companyName = companyName;
            this.recipientType = recipientType;
            this.jobTitle = jobTitle;
            this.emailAddress = "";
        }
    }
}
