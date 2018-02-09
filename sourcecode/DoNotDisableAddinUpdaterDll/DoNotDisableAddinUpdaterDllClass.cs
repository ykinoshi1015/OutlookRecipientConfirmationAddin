using System;
using System.Diagnostics;
using System.Windows.Forms;
using Microsoft.Win32;

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

        /// <summary>
        /// メール内のリンクをクリックしたときにでるセキュリティポップアップを抑制する
        /// </summary>
        /// <param name="protocolName">抑制したいポップアップ(例："notes:")</param>
        /// <returns>抑制設定した場合はtrue、何も設定しなかった場合はfalseを返す</returns>
        public static bool DisableProtocolSecurityPopup(string protocolName)
        {
            bool changed = false;
            string baseRegDirectory = @"Software\Policies\Microsoft\Office\16.0\Common\Security\Trusted Protocols\All Applications";

            //ベースのレジストリキーが存在するか(Officeがインストールされているか)確認する
            try
            {
                RegistryKey regkey = Registry.CurrentUser.OpenSubKey(baseRegDirectory);
                regkey.Close();
            }
            catch (Exception)
            {
                //Officeがインストールされていなければ何もしない
                return changed;
            }

            //すでにDisable設定されているか確認する
            string targetRegDirectory = string.Format("{0}\\{1}", baseRegDirectory, protocolName);
            try
            {
                RegistryKey regkey = Registry.CurrentUser.OpenSubKey(targetRegDirectory);
                regkey.Close();
            }
            catch (NullReferenceException)
            {
                //Disable設定されていな(キーが存在しな)ければ設定する(キーを作る)
                RegistryKey regkey = Registry.CurrentUser.CreateSubKey(targetRegDirectory);
                regkey.Close();
                changed = true;
            }
            catch (Exception)
            {
                ; //予期せぬエラーの時、何もしない
            }

            return changed;
        }
    }
}
