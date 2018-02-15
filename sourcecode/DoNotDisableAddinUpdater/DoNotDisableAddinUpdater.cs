using Microsoft.Win32;
using System;
using System.Collections;
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
        /// <summary>
        /// レジストリキーを確認し、doNotDisable(アドイン無効化の監視をする/しない)と違う値の場合、addinNameの値を設定変更する
        /// </summary>
        /// <param name="addinName">アドイン名</param>
        /// <param name="doNotDisable">アドイン無効化の監視をしないようにするか</param>
        /// <returns>キーの設定変更した場合、true</returns>
        public static bool UpdateDoNotDisableAddinList(string addinName, bool doNotDisable)
        {
            Console.WriteLine("in: DoNotDisableAddinUpdater");

            //レジストリの設定変更したか
            bool updateStatus = false;

            //Outlookのバージョンとレジストリキーのディレクトリを管理するハッシュテーブル
            Hashtable outlookVersionAndDirectory = new Hashtable();

            outlookVersionAndDirectory.Add("Outlook2013", @"Software\Microsoft\Office\15.0\Outlook\Resiliency\DoNotDisableAddinList");
            outlookVersionAndDirectory.Add("Outlook2016", @"Software\Microsoft\Office\16.0\Outlook\Resiliency\DoNotDisableAddinList");

            //見つかったレジストリキーを入れるリスト
            List<RegistryKey> regKeyList = new List<RegistryKey>();

            RegistryKey regKeyTemp = null;

            //Outlookのバージョン別にレジストリキーを取得
            foreach (string directory in outlookVersionAndDirectory.Values)
            {
                regKeyTemp = getOutlookDoNotDisableAddinListRegistryKey(directory);
                
                if (regKeyTemp != null)
                {
                    regKeyList.Add(regKeyTemp);
                }
            }

            //レジストリキーの値がdoNotDisableと違う場合、値をdoNotDisableに変更する
            foreach (RegistryKey regKey in regKeyList)
            {
                Console.WriteLine(doNotDisable);
                Console.WriteLine(Convert.ToInt32(doNotDisable));
                Console.WriteLine(regKey.GetValue(addinName));

                //無効化の監視対象に入っている場合
                if (!Convert.ToInt32(doNotDisable).Equals(regKey.GetValue(addinName)))
                {
                    //キーに、名前/値ペアを書き込む
                    regKey.SetValue(addinName, Convert.ToInt32(doNotDisable), RegistryValueKind.DWord);
                    updateStatus = true;

                }
                //開いたレジストリキーを閉じる
                regKey.Close();
            }

            return updateStatus;
        }
        
        /// <summary>
        /// アドインを無効化にしないレジストリキーを取得する
        /// </summary>
        /// <param name="directory">ハッシュマップの値</param>
        /// <returns>アドイン無効化にしないレジストリキーObject、レジストリキーが存在しない場合null</returns>
        private static RegistryKey getOutlookDoNotDisableAddinListRegistryKey(string directory)
        {
            RegistryKey outlookDoNotDisableAddinListRegistryKey = null;

            try
            {
                //指定したパスのレジストリキーを開く
                outlookDoNotDisableAddinListRegistryKey = Registry.CurrentUser.OpenSubKey(directory, true);
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
