﻿using System;
using NUnit.Framework;
using NSubstitute;
using Microsoft.Office.Interop.Outlook;
using OutlookRecipientConfirmationAddin;
using System.Reflection;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace ORCAUnitTest
{
    /// <summary>
    /// Office365ContactクラスgetContactItemメソッドのテストクラス
    /// </summary>
    [TestFixture]
    public class GetContactItemUnitTest
    {
        private Recipient testRec;
        private Recipient testRec2;
        private AddressEntry testAdd;
        private ExchangeUser testExchUser;
        private Module mod;
        private Type typeThisAddIn;
        private Application testApp;
        private object obj;
        private MethodInfo mi;
        private NameSpace testNs;

        /// <summary>
        /// テスト時に、一度だけ実行される処理（アセンブリの読み込み、Typeの取得など）
        /// </summary>
        [OneTimeSetUp]
        public void Init()
        {
            // テスト対象のメソッド(getContactItem(Recipient recipient)メソッド)の引数のモック
            testRec = Substitute.For<Recipient>();

            // テスト対象のクラス内で使われる変数のモック
            testAdd = Substitute.For<AddressEntry>();
            testExchUser = Substitute.For<ExchangeUser>();
            testRec2 = Substitute.For<Recipient>();

            // ThisAddInクラスを作るのに必要な引数（Factory、IServiceProvider）を作成
            // Factoryクラスは、自作のTestFactoryクラス＆その中で使うTestAddInクラスがないとうまくいかない
            TestFactory testFactory = new TestFactory();
            IServiceProvider testService = Substitute.For<IServiceProvider>();

            // ThisAddInクラスのタイプを取得
            ThisAddIn testAddIn = new ThisAddIn(testFactory, testService);
            typeThisAddIn = testAddIn.GetType();

            // Applicaitionのモック作成
            FieldInfo fieldApp = typeThisAddIn.GetField("Application", BindingFlags.NonPublic | BindingFlags.Instance);
            testApp = Substitute.For<TestApplication>();
            fieldApp.SetValue(testAddIn, testApp);

            // Sessionのモック作成
            testNs = Substitute.For<NameSpace>();
            testNs.CreateRecipient(Arg.Any<string>()).Returns(testRec);
            testApp.Session.Returns(testNs);

            // リフレクション
            // アセンブリを読み込み、モジュールを取得
            //(VSでテストする時)
            //Assembly asm = Assembly.LoadFrom(@".\ORCAUnitTest\bin\Debug\OutlookRecipientConfirmationAddin.dll");
            //(batで実行するとき)
            Assembly asm = Assembly.LoadFrom(@".\OutlookRecipientConfirmationAddin.dll");
            mod = asm.GetModule("OutlookRecipientConfirmationAddin.dll");

            // Globalsのタイプと、ThisAddInプロパティを取得
            Type typeGlobal = mod.GetType("OutlookRecipientConfirmationAddin.Globals");
            PropertyInfo testProp = typeGlobal.GetProperty("ThisAddIn", BindingFlags.NonPublic | BindingFlags.Static);
            // ThisAddinプロパティに、モックなどを使って作った値をセットする
            testProp.SetValue(null, testAddIn);

            // テスト対象のクラス（O365）のタイプを取得
            Type type = mod.GetType("OutlookRecipientConfirmationAddin.Office365Contact");
            // インスタンスを生成し、メソッドにアクセスできるようにする
            obj = Activator.CreateInstance(type);
            mi = type.GetMethod("getContactItem", new Type[] { typeof(Recipient) });
        }

        /// <summary>
        /// RecipientがNotesメールのグループアドレスの場合
        /// ContactItemがnull
        /// </summary>
        [Test]
        public void GetContactItemTest1()
        {
            // テスト用引数の、AddressEntryプロパティとAddressプロパティが呼ばれた場合の返り値を指定
            testRec.AddressEntry.Returns(testAdd);
            testRec.Address.Returns("ZJRITS_ZORG_BS_SKAIC_1KAIB_1G@jrits.ricoh.co.jp");

            // テスト用引数のAddressEntryの、AddressEntryUserTypeプロパティとGetExchangeUserメソッドが呼ばれた場合の返り値を指定
            testAdd.AddressEntryUserType.Returns(OlAddressEntryUserType.olSmtpAddressEntry);
            testAdd.GetExchangeUser().Returns((ExchangeUser)null);

            // テストするメソッドにアクセスし、実際の結果を取得する
            ContactItem actual = (ContactItem)mi.Invoke(obj, new object[] { testRec });

            // nullが返ってくることを確認
            Assert.Null(actual);
        }

        /// <summary>
        /// Recipientがアドレス帳のAll Groupsにあるグループアドレスの場合
        /// ContactItemのFullNameにグループ名が入る
        /// </summary>
        [Test]
        public void GetContactItemTest2()
        {
            string testGroupName = "RITS ビジネスソリューションズ事業部 システム開発センター 第１開発部 第１グループ";

            // テスト用引数の、AddressEntryプロパティとNameプロパティが呼ばれた場合の返り値を指定
            testRec.AddressEntry.Returns(testAdd);
            testRec.Name.Returns(testGroupName);


            // テスト用引数のAddressEntryの、AddressEntryUserTypeプロパティとGetExchangeUserメソッドが呼ばれた場合の返り値を指定
            testAdd.AddressEntryUserType.Returns(OlAddressEntryUserType.olExchangeDistributionListAddressEntry);
            testAdd.GetExchangeUser().Returns(testExchUser);

            // ↓なくてもうまくいくの、いいの？？？
            //ContactItem testContact = Substitute.For<ContactItem>();
            //(testApp as TestApplication).CreateItemHon(Arg.Is(OlItemType.olContactItem)).Returns(testContact);

            // テストするメソッドにアクセスし、実際の結果を取得する
            ContactItem actual = (ContactItem)mi.Invoke(obj, new object[] { testRec });

            Assert.That(actual.FullName, Does.Match(testGroupName));
        }

        /// <summary>
        /// Recipientがアドレス帳の、グローバルアドレス一覧にあるNotesメールのアドレスの場合（グループアドレスでない）
        /// ContactItemのFullName, CompanyName, Department, JobTitleに正しい情報が入る
        /// </summary>
        [Test]
        public void GetContactItemTest3()
        {
            string testName = "Kenta Kosaka/R/RSI";
            string testCompanyName = "Ricoh IT Solutions Co.,Ltd.";
            string testDepartment = "ビジネスソリューションズ事業部 システム開発センター 第１開発部 第１グループ";
            string jobTitle = null;

            // テスト用引数の、AddressEntryプロパティとAddressプロパティが呼ばれた場合の返り値を指定
            testRec.AddressEntry.Returns(testAdd);

            // テスト用引数のAddressEntryの、AddressEntryUserTypeプロパティとGetExchangeUserメソッドが呼ばれた場合の返り値を指定
            testAdd.AddressEntryUserType.Returns(OlAddressEntryUserType.olExchangeRemoteUserAddressEntry);
            testAdd.GetExchangeUser().Returns(testExchUser);

            testExchUser.Name.Returns(testName);
            testExchUser.CompanyName.Returns(testCompanyName);
            testExchUser.Department.Returns(testDepartment);
            testExchUser.JobTitle.Returns(jobTitle);

            // テストするメソッドにアクセスし、実際の結果を取得する
            ContactItem actual = (ContactItem)mi.Invoke(obj, new object[] { testRec });

            Assert.That(actual.FullName, Does.Match(testName));
            Assert.That(actual.CompanyName, Does.Match(testCompanyName));
            Assert.That(actual.Department, Does.Match(testDepartment));
            Assert.Null(actual.JobTitle);
        }

        /// <summary>
        /// Recipientがアドレス帳の、グローバルアドレス一覧にあるOutlookメールのアドレスの場合（グループアドレスでない）
        /// ContactItemのFullName, CompanyName, Department, JobTitleに正しい情報が入る
        /// </summary>
        [Test]
        public void GetContactItemTest4()
        {
            string testName = "Kosaka Kenta (小坂 健太)";
            string testCompanyName = "リコーITソリューションズ";
            string testDepartment = "ビジネスソリューションズ事業部 システム開発センター 第１開発部 第１グループ";
            string testJobTitle = "担当";

            // テスト用引数の、AddressEntryプロパティとAddressプロパティが呼ばれた場合の返り値を指定
            testRec.AddressEntry.Returns(testAdd);

            // テスト用引数のAddressEntryの、AddressEntryUserTypeプロパティとGetExchangeUserメソッドが呼ばれた場合の返り値を指定
            testAdd.AddressEntryUserType.Returns(OlAddressEntryUserType.olExchangeUserAddressEntry);
            testAdd.GetExchangeUser().Returns(testExchUser);

            testExchUser.Name.Returns(testName);
            testExchUser.CompanyName.Returns(testCompanyName);
            testExchUser.Department.Returns(testDepartment);
            testExchUser.JobTitle.Returns(testJobTitle);

            // テストするメソッドにアクセスし、実際の結果を取得する
            ContactItem actual = (ContactItem)mi.Invoke(obj, new object[] { testRec });

            Assert.That(actual.FullName, Does.Contain(testName));
            Assert.That(actual.CompanyName, Does.Match(testCompanyName));
            Assert.That(actual.Department, Does.Match(testDepartment));
            Assert.That(actual.JobTitle, Does.Match(testJobTitle));
        }

        /// <summary>
        /// Recipientが連絡先に登録されたアドレスの場合（グローバルアドレス一覧に存在しないもの）
        /// ContactItemがnull
        /// </summary>
        [Test]
        public void GetContactItemTest5()
        {
            string testAddress = "yna.nakanishi@jp.ricoh.com";

            // テスト用引数の、AddressEntryプロパティとAddressプロパティが呼ばれた場合の返り値を指定
            testRec.AddressEntry.Returns(testAdd);
            testRec.Address.Returns(testAddress);

            // テスト用引数のAddressEntryの、AddressEntryUserTypeプロパティとGetExchangeUserメソッドが呼ばれた場合の返り値を指定
            testAdd.AddressEntryUserType.Returns(OlAddressEntryUserType.olOutlookContactAddressEntry);
            testAdd.GetExchangeUser().Returns((ExchangeUser)null);

            // テストするメソッドにアクセスし、実際の結果を取得する
            ContactItem actual = (ContactItem)mi.Invoke(obj, new object[] { testRec });

            Assert.Null(actual);
        }

        /// <summary>
        /// Recipientが連絡先に登録されたアドレスの場合（Notesメールアドレス）
        /// ContactItemのFullName, CompanyName, Departmentに正しい情報が入る
        /// </summary>
        [Test]
        public void GetContactItemTest6()
        {
            string testAddress = "yasuyuki.kinoshita@jrits.ricoh.co.jp";

            string testName = "Yasuyuki Kinoshita / R / RSI";
            string testCompanyName = "Yasuyuki Kinoshita / R / RSI";
            string testDepartment = "ビジネスソリューションズ事業部 システム開発センター 第１開発部 第１グループ";
            string testJobTitle = null;

            // テスト用引数の、AddressEntryプロパティとAddressプロパティが呼ばれた場合の返り値を指定
            testRec.AddressEntry.Returns(testAdd);
            testRec.Address.Returns(testAddress);

            // テスト用引数のAddressEntryの、AddressEntryUserTypeプロパティとGetExchangeUserメソッドが呼ばれた場合の返り値を指定
            testAdd.AddressEntryUserType.Returns(OlAddressEntryUserType.olOutlookContactAddressEntry);
            testAdd.GetExchangeUser().Returns((ExchangeUser)null);

            // recipient.Addressプロパティを使うと、ExchangeUserの取得に成功する
            testNs.CreateRecipient(testAddress).Returns(testRec2);
            testRec2.AddressEntry.GetExchangeUser().Returns(testExchUser);

            testExchUser.Name.Returns(testName);
            testExchUser.CompanyName.Returns(testCompanyName);
            testExchUser.Department.Returns(testDepartment);
            testExchUser.JobTitle.Returns(testJobTitle);

            //ContactItem testContact = Substitute.For<ContactItem>();
            //(testApp as TestApplication).CreateItemHon(Arg.Is(OlItemType.olContactItem)).Returns(testContact);

            // テストするメソッドにアクセスし、実際の結果を取得する
            ContactItem actual = (ContactItem)mi.Invoke(obj, new object[] { testRec });

            Assert.That(actual.FullName, Does.Contain(testName));
            Assert.That(actual.CompanyName, Does.Match(testCompanyName));
            Assert.That(actual.Department, Does.Match(testDepartment));
            Assert.Null(actual.JobTitle);
        }

        /// <summary>
        /// Recipientが連絡先に登録されたアドレスの場合（O365メールアドレス）
        /// ContactItemのFullName, CompanyName, Department, JobTitleに正しい情報が入る
        /// </summary>
        [Test]
        public void GetContactItemTest7()
        {
            string testAddress = "yasuyuki.kinoshita@jp.ricoh.com";

            string testName = "Kinoshita Yasuyuki (木下 康行)";
            string testCompanyName = "リコーITソリューションズ";
            string testDepartment = "ビジネスソリューションズ事業部 システム開発センター 第１開発部 第１グループ";
            string testJobTitle = "担当";

            // テスト用引数の、AddressEntryプロパティとAddressプロパティが呼ばれた場合の返り値を指定
            testRec.AddressEntry.Returns(testAdd);
            testRec.Address.Returns(testAddress);

            // テスト用引数のAddressEntryの、AddressEntryUserTypeプロパティとGetExchangeUserメソッドが呼ばれた場合の返り値を指定
            testAdd.AddressEntryUserType.Returns(OlAddressEntryUserType.olOutlookContactAddressEntry);
            testAdd.GetExchangeUser().Returns((ExchangeUser)null);

            // recipient.Addressプロパティを使うと、ExchangeUserの取得に成功する
            testNs.CreateRecipient(testAddress).Returns(testRec2);
            testRec2.AddressEntry.GetExchangeUser().Returns(testExchUser);

            testExchUser.Name.Returns(testName);
            testExchUser.CompanyName.Returns(testCompanyName);
            testExchUser.Department.Returns(testDepartment);
            testExchUser.JobTitle.Returns(testJobTitle);

            // テストするメソッドにアクセスし、実際の結果を取得
            ContactItem actual = (ContactItem)mi.Invoke(obj, new object[] { testRec });

            Assert.That(actual.FullName, Does.Contain(testName));
            Assert.That(actual.CompanyName, Does.Match(testCompanyName));
            Assert.That(actual.Department, Does.Match(testDepartment));
            Assert.That(actual.JobTitle, Does.Match(testJobTitle));
        }

    }

}