using System;
using System.Diagnostics;
using System.Windows.Forms;

namespace DoNotDisableAddinUpdaterDll
{
    public class DoNotDisableAddinUpdaterDllClass
    {
        /// <summary>
        /// レジストリを確認し、無効化しない設定でない場合、キーを追加し設定する
        /// </summary>
        /// <param name="version">起動したOutlookのバージョン</param>
        public static void checkDisable(string version)
        {
            Debug.WriteLine("in: DoNotDisableAddinUpdaterDllClass");

            string regkeyDirectory = null;
            Microsoft.Win32.RegistryKey regkey = null;

            try
            {
                //Outlook2016を使っている場合
                if (version.StartsWith("16."))
                {
                    //アドイン無効化の監視対象の設定をするキーを開く
                    regkeyDirectory = @"Software\Microsoft\Office\16.0\Outlook\Resiliency\DoNotDisableAddinList";
                    regkey = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(regkeyDirectory, true);
                }
                //Outlook2013を使っている場合
                else if (version.StartsWith("15."))
                {
                    regkeyDirectory = @"Software\Microsoft\Office\15.0\Outlook\Resiliency\DoNotDisableAddinList";
                    regkey = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(regkeyDirectory, true);
                }
                else
                {
                    MessageBox.Show("お使いのOutlookのバージョンでは、Outlook宛先表示アドインをご利用になれません。");
                }

                ///キーの名前が見つからない場合はnullが帰る
                object regkeyValue = regkey.GetValue("OutlookRecipientConfirmationAddin");

                //無効化しない設定がされていない場合（キー、名前/値ペアを設定する）
                //キーの名前が存在しない場合(作成する)
                if (!0x000000001.Equals(regkeyValue))
                {
                    //REG_DWORDで書き込む
                    regkey.SetValue("OutlookRecipientConfirmationAddin", 0x000000001, Microsoft.Win32.RegistryValueKind.DWord);
                }

            }
            //キーが存在しない場合
            //新規作成で開いて、値を設定する
            catch (NullReferenceException)
            {
                regkey = Microsoft.Win32.Registry.CurrentUser.CreateSubKey(regkeyDirectory);
                regkey.SetValue("OutlookRecipientConfirmationAddin", 0x000000001, Microsoft.Win32.RegistryValueKind.DWord);
            }

            //閉じる
            regkey.Close();
        }
    }
}
