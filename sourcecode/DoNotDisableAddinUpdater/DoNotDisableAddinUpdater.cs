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
            Debug.WriteLine("in: DoNotDisableAddinUpdater");

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

                    if (regKeyTemp != null)
                    {
                        regKeyList.Add(regKeyTemp);
                    }
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
                //Console.WriteLine(doNotDisable);
                //Console.WriteLine(Convert.ToInt32(doNotDisable));
                //Console.WriteLine(regKey.GetValue(addinName));

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

        ///// <summary>
        ///// メール内のリンクをクリックしたときにでるセキュリティポップアップを抑制する
        ///// </summary>
        ///// <param name="protocolName">抑制したいポップアップ(例："notes:")</param>
        ///// <returns>抑制設定した場合はtrue、何も設定しなかった場合はfalseを返す</returns>
        //public static bool DisableProtocolSecurityPopup(string protocolName)
        //{
        //    bool changed = false;
        //    string baseRegDirectory = @"Software\Policies\Microsoft\Office\16.0\Common\Security\Trusted Protocols\All Applications";

        //    //ベースのレジストリキーが存在するか(Officeがインストールされているか)確認する
        //    try
        //    {
        //        RegistryKey regkey = Registry.CurrentUser.OpenSubKey(baseRegDirectory);
        //        regkey.Close();
        //    }
        //    catch (Exception)
        //    {
        //        //Officeがインストールされていなければ何もしない
        //        return changed;
        //    }

        //    //すでにDisable設定されているか確認する
        //    string targetRegDirectory = string.Format("{0}\\{1}", baseRegDirectory, protocolName);
        //    try
        //    {
        //        RegistryKey regkey = Registry.CurrentUser.OpenSubKey(targetRegDirectory);
        //        regkey.Close();
        //    }
        //    catch (NullReferenceException)
        //    {
        //        //Disable設定されていな(キーが存在しな)ければ設定する(キーを作る)
        //        RegistryKey regkey = Registry.CurrentUser.CreateSubKey(targetRegDirectory);
        //        regkey.Close();
        //        changed = true;
        //    }
        //    catch (Exception)
        //    {
        //        ; //予期せぬエラーの時、何もしない
        //    }

        //    return changed;
        //}

    }
}
