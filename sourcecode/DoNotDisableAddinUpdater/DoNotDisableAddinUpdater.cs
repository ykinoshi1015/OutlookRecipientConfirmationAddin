using Microsoft.Win32;
using System;
using System.Diagnostics;
using System.Linq;
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

            RegistryKey regKey = null;
            RegistryKey regKeyTemp = null;

            //Outlookのバージョン別にレジストリキーを取得
            foreach (OutlookVersion outlookversion in Enum.GetValues(typeof(OutlookVersion)))
            {
                regKeyTemp = getOutlookDoNotDisableAddinListRegistryKey(outlookversion);

                if (regKeyTemp != null)
                {
                    regKey = regKeyTemp;
                }
            }

            //キーが存在しない場合、設定を変更しない
            if (regKey == null)
            {
                return updateStatus;
            }

            //アドイン名がOutlookRecipientConfirmationAddin、無効化の監視対象を外したい場合
            //かつ、無効化の監視対象に入っている場合
            if (addinName.Equals("OutlookRecipientConfirmationAddin") && doNotDisable && !0x00000001.Equals(regKey.GetValue(addinName)))
            {
                //キー、名前/値ペアを書き込む
                regKey.SetValue(addinName, 0x000000001, RegistryValueKind.DWord);

                updateStatus = true;
            }

            //開いたレジストリキーを閉じる
            regKey.Close();

            return updateStatus;
        }


        /// <summary>
        /// アドインを無効化にしないレジストリキーを取得する
        /// </summary>
        /// <param name="version">起動したOutlookのバージョン</param>
        /// <returns>アドイン無効化にしないレジストリキーObject</returns>
        private static RegistryKey getOutlookDoNotDisableAddinListRegistryKey(OutlookVersion version)
        {
            RegistryKey outlookDoNotDisableAddinListRegistryKey = null;
            string regkeyDirectory = null;

            try
            {
                //操作するレジストリ・キーの名前を取得
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
                outlookDoNotDisableAddinListRegistryKey = Registry.CurrentUser.OpenSubKey(regkeyDirectory, true);

            }
            //OpenSubKeyに失敗した場合
            catch (NullReferenceException)
            {
                return null;
            }

            //RegistryKeyオブジェクトが見つかったら返す
            return outlookDoNotDisableAddinListRegistryKey;

        }
    }
}
