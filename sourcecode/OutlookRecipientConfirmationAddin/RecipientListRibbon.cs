using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;
using Outlook = Microsoft.Office.Interop.Outlook;

// TODO:  リボン (XML) アイテムを有効にするには、次の手順に従います。


// 2. ボタンのクリックなど、ユーザーの操作を処理するためのコールバック メソッドを、このクラスの
//    "リボンのコールバック" 領域に作成します。メモ: このリボンがリボン デザイナーからエクスポートされたものである場合は、
//    イベント ハンドラー内のコードをコールバック メソッドに移動し、リボン拡張機能 (RibbonX) のプログラミング モデルで
//    動作するように、コードを変更します。

// 3. リボン XML ファイルのコントロール タグに、コードで適切なコールバック メソッドを識別するための属性を割り当てます。  

// 詳細については、Visual Studio Tools for Office ヘルプにあるリボン XML のドキュメントを参照してください。


namespace OutlookRecipientConfirmationAddin
{
    [ComVisible(true)]
    public class RecipientListRibbon : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;

        public RecipientListRibbon()
        {
        }

        #region IRibbonExtensibility のメンバー

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("OutlookRecipientConfirmationAddin.RecipientListRibbon.xml");
        }

        #endregion

        #region リボンのコールバック
        //ここにコールバック メソッドを作成します。コールバック メソッドの追加方法の詳細については、http://go.microsoft.com/fwlink/?LinkID=271226 にアクセスしてください。

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;

        }

        /// <summary>
        /// 宛先確認ボタンが押された場合
        /// この中で、そのメールの受信者の一覧を探してきて、次の画面に渡す？
        /// </summary>
        /// <param name="ribbonUI"></param>
        public void RecipientListButton_Click(Office.IRibbonControl ribbonUI)
        {
            /// MessageBox.Show("宛先確認ボタンがおされました");
            try
            {

                Microsoft.Office.Interop.Outlook.NameSpace objNamespace = Globals.ThisAddIn.Application.GetNamespace("MAPI");

                var explorer = Globals.ThisAddIn.Application.ActiveExplorer();
                Microsoft.Office.Interop.Outlook.Selection selection = explorer.Selection;
                object selectedItem = selection[1];

                ///string str = mailItem.To;
               
                ///とりあえずmailにしてみた
                RecipientConfirmationWindow.SendType type = RecipientConfirmationWindow.SendType.Mail;


                Outlook.Recipients recipients = null;
                Outlook.MailItem mail = selectedItem as Outlook.MailItem;

                if (selectedItem != null)
                {
                    recipients = mail.Recipients;
                    type = RecipientConfirmationWindow.SendType.Mail;
                }
                else
                {
                    Outlook.MeetingItem meeting = selectedItem as Outlook.MeetingItem;
                    if (meeting != null)
                    {
                        recipients = meeting.Recipients;
                        type = RecipientConfirmationWindow.SendType.Meeting;
                    }
                }

                /// 受信者の情報をリストする
                List<Outlook.Recipient> recipientsList = new List<Outlook.Recipient>();
                foreach (Outlook.Recipient recipient in recipients)
                {
                    recipientsList.Add(recipient);
                }

                /// 検索クラスを呼び出す
                SearchRecipient searchRecipient = new SearchRecipient();

                /// 引数に宛先に指定されたアドレスのリストを渡すと、宛先情報のリストが戻ってくる
                List<RecipientInformationDto> recipientList = searchRecipient.SearchContact(recipientsList);


                /// 宛先リストの画面を表示する
                RecipientListWindow recipientListWindow = new RecipientListWindow(type, recipientList);
                DialogResult result = recipientListWindow.ShowDialog();


            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
            }
        }



        #endregion

        #region ヘルパー

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }

        #endregion
    }
}
