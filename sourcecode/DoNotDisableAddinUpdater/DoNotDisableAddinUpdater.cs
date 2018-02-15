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
        private static readonly string[] _doNotDisableAddinListKeys ={
            @"Software\Microsoft\Office\15.0\Outlook\Resiliency\DoNotDisableAddinList",
            @"Software\Microsoft\Office\16.0\Outlook\Resiliency\DoNotDisableAddinList",
        };

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

            //見つかったレジストリキーを入れるリスト
            List<RegistryKey> regKeyList = new List<RegistryKey>();

            //Outlookのバージョン別にレジストリキーを取得
            foreach (string regKey in _doNotDisableAddinListKeys)
            {
                RegistryKey regKeyTemp = null;

                try
                {
                    //指定したパスのレジストリキーを開く
                    regKeyTemp = Registry.CurrentUser.OpenSubKey(regKey, true);
                    regKeyList.Add(regKeyTemp);
                }
                //OpenSubKeyに失敗した場合
                catch (NullReferenceException)
                {
                    continue;
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

    }
}
