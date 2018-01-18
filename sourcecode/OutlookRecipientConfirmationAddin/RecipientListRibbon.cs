using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;

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
        /// </summary>
        /// <param name="ribbonUI"></param>
        public void RecipientListButton_Click(Office.IRibbonControl ribbonUI)
        {
            /// MessageBox.Show("宛先確認ボタンがおされました");

            /// 宛先リストの画面を表示する
            RecipientListWindow recipientListWindow = new RecipientListWindow();
            DialogResult result = recipientListWindow.ShowDialog();
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
