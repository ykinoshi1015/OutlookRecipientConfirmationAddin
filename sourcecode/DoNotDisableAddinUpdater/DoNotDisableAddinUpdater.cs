using Microsoft.Win32;
using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace DoNotDisableAddinUpdater
{
    /// <summary>
    /// アドイン無効化の設定を更新するクラス
    /// </summary>
    public class DoNotDisableAddinListUpdater
    {
        //起動したOutlookのバージョン
        enum OutlookVersion { Outlook2013, Outlook2016 };

        /// <summary>
        /// レジストリを確認し、無効化しない設定でない場合、キーを追加し設定する
        /// </summary>
        /// <param name="addinName">アドイン名</param>
        /// <param name="doNotDisable">アドイン無効化の監視をしないようにするか</param>
        /// <returns>レジストリの設定変更した場合、true</returns>
        public static bool UpdateDoNotDisableAddinList(string addinName, bool doNotDisable)
        {
            Debug.WriteLine("in: DoNotDisableAddinUpdater");

            //レジストリの設定変更したか
            bool updateStatus = false;

            //起動中のOutlookのインスタンスを取得
            Outlook.Application application = Marshal.GetActiveObject("Outlook.Application") as Outlook.Application;
            OutlookVersion outlookVersion = OutlookVersion.Outlook2013;

            if (application.Version.StartsWith("15."))
            {
                outlookVersion = OutlookVersion.Outlook2013;
            }
            else if (application.Version.StartsWith("16."))
            {
                outlookVersion = OutlookVersion.Outlook2016;
            }

            RegistryKey outlookDoNotDisableAddinListRegistryKey = getOutlookDoNotDisableAddinListRegistryKey(outlookVersion);

            //キーまたは値が存在しない場合、設定を変更しない
            if (outlookDoNotDisableAddinListRegistryKey == null)
            {
                return updateStatus;
            }

            //アドイン名がOutlookRecipientConfirmationAddinで、アドイン無効化の監視対象にする場合
            if (doNotDisable && addinName.Equals("OutlookRecipientConfirmationAddin"))
            {
                //無効化しない設定がされていない場合（キー、名前/値ペアを設定する）
                if (!0x000000001.Equals(outlookDoNotDisableAddinListRegistryKey))
                {
                    //REG_DWORDで書き込む
                    outlookDoNotDisableAddinListRegistryKey.SetValue(addinName, 0x000000001, RegistryValueKind.DWord);
                    updateStatus = true;
                }

            }
            return updateStatus;

        }


        /// <summary>
        /// Outlookの無効化にしないレジストリキーを取得する
        /// </summary>
        /// <param name="version">起動したOutlookのバージョン</param>
        /// <returns>Outlookの無効化にしないレジストリキーObject</returns>
        private static RegistryKey getOutlookDoNotDisableAddinListRegistryKey(OutlookVersion version)
        {
            RegistryKey outlookDoNotDisableAddinListRegistryKey = null;
            string regkeyDirectory = null;
            RegistryKey regkey = null;

            try
            {
                //アドイン無効化の監視対象の設定をするキーを開く
                switch (version)
                {
                    case OutlookVersion.Outlook2013:
                        regkeyDirectory = @"Software\Microsoft\Office\15.0\Outlook\Resiliency\DoNotDisableAddinList";
                        break;

                    case OutlookVersion.Outlook2016:
                        regkeyDirectory = @"Software\Microsoft\Office\16.0\Outlook\Resiliency\DoNotDisableAddinList";
                        break;
                }

                //指定したパスのレジストリキーを開く
                regkey = Registry.CurrentUser.OpenSubKey(regkeyDirectory, true);

                //レジストリの値を取得
                outlookDoNotDisableAddinListRegistryKey = regkey.GetValue("OutlookRecipientConfirmationAddin") as RegistryKey;
            }
            //キーまたは値が存在しない場合、nullを返す
            catch (NullReferenceException)
            {
                return null;
            }
            //開いたレジストリキーを閉じる
            finally
            {
                regkey.Close();
            }

            //キーが見つかった場合、RegistryKeyオブジェクトを返す
            return outlookDoNotDisableAddinListRegistryKey;

        }
    }
}
