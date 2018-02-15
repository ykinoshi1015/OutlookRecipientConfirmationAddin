using Microsoft.Win32;
using System;
using System.Collections.Generic;
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
        private enum OutlookVersion { Outlook2013, Outlook2016 };

        /// <summary>
        /// レジストリを確認し、無効化しない設定でない場合、キーを追加し設定する
        /// </summary>
        /// <param name="addinName">アドイン名</param>
        /// <param name="doNotDisable">アドイン無効化の監視をしないようにするか</param>
        /// <returns>レジストリの設定変更した場合、true</returns>
        public static bool UpdateDoNotDisableAddinList(string addinName, bool doNotDisable)
        {
            Console.WriteLine("in: DoNotDisableAddinUpdater");

            //レジストリの設定変更したか
            bool updateStatus = false;

            List<RegistryKey> regKeyList = new List<RegistryKey>();
            RegistryKey regKeyTemp = null;

            //Outlookのバージョン別にレジストリキーを取得
            foreach (OutlookVersion outlookversion in Enum.GetValues(typeof(OutlookVersion)))
            {
                regKeyTemp = getOutlookDoNotDisableAddinListRegistryKey(outlookversion);

                if (regKeyTemp != null)
                {
                    regKeyList.Add(regKeyTemp);
                }
            }

            //キーが存在しない場合、設定を変更しない
            if (regKeyList == null)
            {
                return updateStatus;
            }


            foreach (RegistryKey regKey in regKeyList)
            {
                Console.WriteLine(doNotDisable);
                Console.WriteLine(Convert.ToInt32(doNotDisable));
                Console.WriteLine(regKey.GetValue(addinName));

                //無効化の監視対象に入っている場合
                if (!Convert.ToInt32(doNotDisable).Equals(regKey.GetValue(addinName)))
                {
                    //キー、名前/値ペアを書き込み、監視対象外にする
                    regKey.SetValue(addinName, Convert.ToInt32(doNotDisable), RegistryValueKind.DWord);

                    if (!updateStatus)
                    {
                        updateStatus = true;
                    }

                }
                //開いたレジストリキーを閉じる
                regKey.Close();
            }

            return updateStatus;
        }


        /// <summary>
        /// アドインを無効化にしないレジストリキーを取得する
        /// </summary>
        /// <param name="version">起動したOutlookのバージョン</param>
        /// <returns>アドイン無効化にしないレジストリキーObject、キーが存在しない場合null</returns>
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
