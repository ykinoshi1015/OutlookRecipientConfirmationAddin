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
        List<String> toList;
        List<String> ccList;
        List<String> bccList;

        /// 検索結果の宛先情報のリスト
        private List<RecipientInformationDto> RecipientInformationList;

        /// コンストラクタ
        public SearchRecipient(List<String> toList, List<String> ccList, List<String> bccList)
        {
            this.toList = toList;
            this.ccList = ccList;
            this.bccList = bccList;
        }

        /// <summary>
        /// メールのアドレスから宛先情報を検索する
        /// </summary>
        /// <param name="addressList"></param> メールのTO, CC, BCC
        /// <returns> 検索した宛先情報のリスト</returns>
        public List<RecipientInformationDto> SearchContact(List<String> addressList)
        {
            /// ファクトリオブジェクトに連絡先クラスのインスタンスの生成をしてもらう
            ContactFactory contactFactory = new ContactFactory();
            c = contactFactory.CreateContact();



            return RecipientInformationList;
        }
    }
}
