using System;
using NUnit.Framework;
using NSubstitute;
using Microsoft.Office.Interop.Outlook;
using OutlookRecipientConfirmationAddin;
using System.Reflection;
using Microsoft.Office.Tools;
using Microsoft.Office.Tools.Outlook;
using Microsoft.Office.Tools.Ribbon;
using System.Collections.Generic;
using System.ComponentModel;
using System.Windows.Forms;
using System.Dynamic;
using Outlook = Microsoft.Office.Interop.Outlook;


namespace ORCAUnitTest
{
    /// <summary>
    /// Outlook宛先表示アドインの主要メソッドテストクラス
    /// </summary>
    [TestFixture]
    public class UnitTest1
    {
        private Recipient testRec;
        private AddressEntry testAdd;
        private ExchangeUser testExchUser;
        private Module mod;
        private Type typeThisAddIn;
        private Microsoft.Office.Interop.Outlook.Application testApp;
        private object obj;
        private MethodInfo mi;
        private NameSpace testNs;

        private object obj2;
        private MethodInfo mi2;
        /// <summary>
        /// 
        /// </summary>
        [OneTimeSetUp]
        public void Init()
        {
            // テスト対象のメソッド(getContactItem(Recipient recipient)メソッド)の引数のモック
            testRec = Substitute.For<Recipient>();

            // テスト対象のクラス内で使われる変数のモック
            testAdd = Substitute.For<AddressEntry>();
            testExchUser = Substitute.For<ExchangeUser>();

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
            Assembly asm = Assembly.LoadFrom(@".\ORCAUnitTest\bin\Debug\OutlookRecipientConfirmationAddin.dll");
            //(batで実行するとき)
            //Assembly asm = Assembly.LoadFrom(@".\OutlookRecipientConfirmationAddin.dll");
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

            // テスト対象のクラス（Utility）のタイプを取得
            Type type2 = mod.GetType("OutlookRecipientConfirmationAddin.Utility");
            // インスタンスを生成し、メソッドにアクセスできるようにする
            obj2 = Activator.CreateInstance(type2);
            // mi2 = type2.GetMethod("GetRecipients", new Type[] { typeof(object), typeof(Utility.OutlookItemType), typeof(bool)  });
            mi2 = type2.GetMethod("GetRecipients");
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
            // テスト用引数の、AddressEntryプロパティとNameプロパティが呼ばれた場合の返り値を指定
            testRec.AddressEntry.Returns(testAdd);
            testRec.Name.Returns("RITS ビジネスソリューションズ事業部 システム開発センター 第１開発部 第１グループ");


            // テスト用引数のAddressEntryの、AddressEntryUserTypeプロパティとGetExchangeUserメソッドが呼ばれた場合の返り値を指定
            testAdd.AddressEntryUserType.Returns(OlAddressEntryUserType.olExchangeDistributionListAddressEntry);
            testAdd.GetExchangeUser().Returns(testExchUser);

            // ↓なくてもうまくいくの、いいの？？？
            //ContactItem testContact = Substitute.For<ContactItem>();
            //(testApp as TestApplication).CreateItemHon(Arg.Is(OlItemType.olContactItem)).Returns(testContact);

            // テストするメソッドにアクセスし、実際の結果を取得する
            ContactItem actual = (ContactItem)mi.Invoke(obj, new object[] { testRec });

            Assert.That(actual.FullName, Does.Match("RITS ビジネスソリューションズ事業部 システム開発センター 第１開発部 第１グループ"));
        }

        /// <summary>
        /// Recipientがアドレス帳の、グローバルアドレス一覧にあるNotesメールのアドレスの場合（グループアドレスでない）
        /// ContactItemのFullName, CompanyName, Department, JobTitleに正しい情報が入る
        /// </summary>
        [Test]
        public void GetContactItemTest3()
        {
            // kenta.kosaka @jrits.ricoh.co.jp

            // テスト用引数の、AddressEntryプロパティとAddressプロパティが呼ばれた場合の返り値を指定
            testRec.AddressEntry.Returns(testAdd);

            // テスト用引数のAddressEntryの、AddressEntryUserTypeプロパティとGetExchangeUserメソッドが呼ばれた場合の返り値を指定
            testAdd.AddressEntryUserType.Returns(OlAddressEntryUserType.olExchangeRemoteUserAddressEntry);
            testAdd.GetExchangeUser().Returns(testExchUser);

            testExchUser.Name.Returns("Kenta Kosaka/R/RSI");
            testExchUser.CompanyName.Returns("Ricoh IT Solutions Co.,Ltd.");
            testExchUser.Department.Returns("ビジネスソリューションズ事業部 システム開発センター 第１開発部 第１グループ");
            testExchUser.JobTitle.Returns((string)null);

            // テストするメソッドにアクセスし、実際の結果を取得する
            ContactItem actual = (ContactItem)mi.Invoke(obj, new object[] { testRec });

            Assert.That(actual.FullName, Does.Match("Kenta Kosaka/R/RSI"));
            Assert.That(actual.CompanyName, Does.Match("Ricoh IT Solutions Co.,Ltd."));
            Assert.That(actual.Department, Does.Match("ビジネスソリューションズ事業部 システム開発センター 第１開発部 第１グループ"));
            Assert.Null(actual.JobTitle);
        }

        /// <summary>
        /// Recipientがアドレス帳の、グローバルアドレス一覧にあるOutlookメールのアドレスの場合（グループアドレスでない）
        /// ContactItemのFullName, CompanyName, Department, JobTitleに正しい情報が入る
        /// </summary>
        [Test]
        public void GetContactItemTest4()
        {
            //kosaka.kenta@jp.ricoh.com

            // テスト用引数の、AddressEntryプロパティとAddressプロパティが呼ばれた場合の返り値を指定
            testRec.AddressEntry.Returns(testAdd);

            // テスト用引数のAddressEntryの、AddressEntryUserTypeプロパティとGetExchangeUserメソッドが呼ばれた場合の返り値を指定
            testAdd.AddressEntryUserType.Returns(OlAddressEntryUserType.olExchangeUserAddressEntry);
            testAdd.GetExchangeUser().Returns(testExchUser);

            testExchUser.Name.Returns("Kosaka Kenta (小坂 健太)");
            testExchUser.CompanyName.Returns("リコーITソリューションズ");
            testExchUser.Department.Returns("ビジネスソリューションズ事業部 システム開発センター 第１開発部 第１グループ");
            testExchUser.JobTitle.Returns("担当");

            // テストするメソッドにアクセスし、実際の結果を取得する
            ContactItem actual = (ContactItem)mi.Invoke(obj, new object[] { testRec });

            Assert.That(actual.FullName, Does.Contain("Kosaka Kenta (小坂 健太)"));
            Assert.That(actual.CompanyName, Does.Match("リコーITソリューションズ"));
            Assert.That(actual.Department, Does.Match("ビジネスソリューションズ事業部 システム開発センター 第１開発部 第１グループ"));
            Assert.That(actual.JobTitle, Does.Match("担当"));
        }

        /// <summary>
        /// Recipientが連絡先に登録されたアドレスの場合（グローバルアドレス一覧に存在しないもの）
        /// ContactItemがnull
        /// </summary>
        [Test]
        public void GetContactItemTest5()
        {
            // テスト用引数の、AddressEntryプロパティとAddressプロパティが呼ばれた場合の返り値を指定
            testRec.AddressEntry.Returns(testAdd);
            testRec.Address.Returns("yna.nakanishi@jp.ricoh.com");

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
            // テスト用引数の、AddressEntryプロパティとAddressプロパティが呼ばれた場合の返り値を指定
            testRec.AddressEntry.Returns(testAdd);
            testRec.Address.Returns("yasuyuki.kinoshita@jrits.ricoh.co.jp");

            // テスト用引数のAddressEntryの、AddressEntryUserTypeプロパティとGetExchangeUserメソッドが呼ばれた場合の返り値を指定
            testAdd.AddressEntryUserType.Returns(OlAddressEntryUserType.olOutlookContactAddressEntry);
            testAdd.GetExchangeUser().Returns((ExchangeUser)null);

            // recipient.Addressプロパティを使うと、ExchangeUserの取得に成功する
            Recipient testRec2 = Substitute.For<Recipient>();
            testNs.CreateRecipient("yasuyuki.kinoshita@jrits.ricoh.co.jp").Returns(testRec2);
            testRec2.AddressEntry.GetExchangeUser().Returns(testExchUser);

            testExchUser.Name.Returns("Yasuyuki Kinoshita / R / RSI");
            testExchUser.CompanyName.Returns("Ricoh IT Solutions Co.,Ltd.");
            testExchUser.Department.Returns("ビジネスソリューションズ事業部 システム開発センター 第１開発部 第１グループ");
            testExchUser.JobTitle.Returns((string)null);

            //ContactItem testContact = Substitute.For<ContactItem>();
            //(testApp as TestApplication).CreateItemHon(Arg.Is(OlItemType.olContactItem)).Returns(testContact);

            // テストするメソッドにアクセスし、実際の結果を取得する
            ContactItem actual = (ContactItem)mi.Invoke(obj, new object[] { testRec });

            Assert.That(actual.FullName, Does.Contain("Yasuyuki Kinoshita / R / RSI"));
            Assert.That(actual.CompanyName, Does.Match("Ricoh IT Solutions Co.,Ltd."));
            Assert.That(actual.Department, Does.Match("ビジネスソリューションズ事業部 システム開発センター 第１開発部 第１グループ"));
            Assert.Null(actual.JobTitle);
        }

        /// <summary>
        /// Recipientが連絡先に登録されたアドレスの場合（O365メールアドレス）
        /// ContactItemのFullName, CompanyName, Department, JobTitleに正しい情報が入る
        /// </summary>
        [Test]
        public void GetContactItemTest7()
        {
            // テスト用引数の、AddressEntryプロパティとAddressプロパティが呼ばれた場合の返り値を指定
            testRec.AddressEntry.Returns(testAdd);
            testRec.Address.Returns("yasuyuki.kinoshita@jp.ricoh.com");

            // テスト用引数のAddressEntryの、AddressEntryUserTypeプロパティとGetExchangeUserメソッドが呼ばれた場合の返り値を指定
            testAdd.AddressEntryUserType.Returns(OlAddressEntryUserType.olOutlookContactAddressEntry);
            testAdd.GetExchangeUser().Returns((ExchangeUser)null);

            // recipient.Addressプロパティを使うと、ExchangeUserの取得に成功する
            Recipient testRec2 = Substitute.For<Recipient>();
            testNs.CreateRecipient("yasuyuki.kinoshita@jp.ricoh.com").Returns(testRec2);
            testRec2.AddressEntry.GetExchangeUser().Returns(testExchUser);

            testExchUser.Name.Returns("Kinoshita Yasuyuki (木下 康行)");
            testExchUser.CompanyName.Returns("リコーITソリューションズ");
            testExchUser.Department.Returns("ビジネスソリューションズ事業部 システム開発センター 第１開発部 第１グループ");
            testExchUser.JobTitle.Returns("担当");


            // テストするメソッドにアクセスし、実際の結果を取得
            ContactItem actual = (ContactItem)mi.Invoke(obj, new object[] { testRec });

            Assert.That(actual.FullName, Does.Contain("Kinoshita Yasuyuki (木下 康行)"));
            Assert.That(actual.CompanyName, Does.Match("リコーITソリューションズ"));
            Assert.That(actual.Department, Does.Match("ビジネスソリューションズ事業部 システム開発センター 第１開発部 第１グループ"));
            Assert.That(actual.JobTitle, Does.Match("担当"));
        }

        /// <summary>
        ///  MailItemの場合
        ///  Recipientsを取得でき、TypeがMailのままになる
        /// </summary>
        [Test]
        public void GetRecipientsTest1()
        {
            // テスト用のMailItemを、モックで作成
            MailItem testMail = Substitute.For<MailItem>();
            
            // モックでつかうデータを用意
            string[] testRecNames = { "testemailaddress1@example.com", "testemailaddress2@example.com" };
            bool[] testRecSendable = { true, true };
            int[] testRecType = { (int)OlMailRecipientType.olTo, (int)OlMailRecipientType.olCC };

            // モックのReturn値を設定
            testMail.Recipients.Count.Returns(testRecNames.Length);

            int i = 0;
            foreach (string testRec in testRecNames)
            {
                testMail.Recipients[i + 1].Address.Returns(testRecNames[i]);
                testMail.Recipients[i + 1].Sendable.Returns(testRecSendable[i]);
                testMail.Recipients[i + 1].Type.Returns(testRecType[i]);
                i++;
            }
            
            // テストするメソッドにアクセスし、実際の結果を取得
            // ここではList<Recipient>にキャストできない（理由は？）
            var objArray = new object[] { testMail, Utility.OutlookItemType.Mail, false };
            object actualObj = mi2.Invoke(obj2, objArray);

            // テスト対象メソッドの返り値をList<Recipient>型にする
            List<Recipient> actualRecList = new List<Recipient>();
            IEnumerable<Recipient> actualEnumList = (IEnumerable<Recipient>)actualObj;

            foreach (var actual in actualEnumList)
            {
                actualRecList.Add(actual);
            }
            
            // 期待結果を入れるリスト
            List<Recipient> expectedRecList = new List<Recipient>();
            
            // 期待結果1のデータをリストに追加
            Recipient expectedRec1 = Substitute.For<Recipient>();
            expectedRec1.Address.Returns("testemailaddress1@example.com");
            expectedRec1.Sendable.Returns(true);
            expectedRec1.Type.Returns((int)OlMailRecipientType.olTo);
            expectedRecList.Add(expectedRec1);

            // 期待結果2のデータをリストに追加
            Recipient expectedRec2 = Substitute.For<Recipient>();
            expectedRec2.Address.Returns("testemailaddress2@example.com");
            expectedRec2.Sendable.Returns(true);
            expectedRec2.Type.Returns((int)OlMailRecipientType.olCC);
            expectedRecList.Add(expectedRec2);

            // actualとexpectedのリストを比較
            Assert.AreEqual(actualRecList.Count, expectedRecList.Count);

            Assert.That(actualRecList[0].Address, Is.EqualTo(expectedRecList[0].Address));
            Assert.That(actualRecList[0].Sendable, Is.EqualTo(expectedRecList[0].Sendable));
            Assert.That(actualRecList[0].Type, Is.EqualTo(expectedRecList[0].Type));

            Assert.That(actualRecList[1].Address, Is.EqualTo(expectedRecList[1].Address));
            Assert.That(actualRecList[1].Sendable, Is.EqualTo(expectedRecList[1].Sendable));
            Assert.That(actualRecList[1].Type, Is.EqualTo(expectedRecList[1].Type));

            // ref引数のtypeが正しいことを確認
            Assert.That(objArray[1], Is.EqualTo(Utility.OutlookItemType.Mail));
            
        }

        /// <summary>
        ///  ItemがMeetingItemの場合
        ///  Recipientsを取得でき、TypeがMeetingItemになる
        /// </summary>
        [Test]
        public void GetRecipientsTest2()
        {
            // テスト用のMailItemを、モックで作成
            MeetingItem testMeeting = Substitute.For<MeetingItem>();

            // モックでつかうデータを用意
            string[] testRecNames = { "testemailaddress1@example.com", "testemailaddress2@example.com" };
            bool[] testRecSendable = { true, true };
            int[] testRecType = { (int)OlMailRecipientType.olTo, (int)OlMailRecipientType.olCC };

            // モックのReturn値を設定
            testMeeting.Recipients.Count.Returns(testRecNames.Length);
            testMeeting.MessageClass.Returns("IPM.Schedule.Meeting.Request");

            int i = 0;
            foreach (string testRec in testRecNames)
            {
                testMeeting.Recipients[i + 1].Address.Returns(testRecNames[i]);
                testMeeting.Recipients[i + 1].Sendable.Returns(testRecSendable[i]);
                testMeeting.Recipients[i + 1].Type.Returns(testRecType[i]);
                i++;
            }

            // テストするメソッドにアクセスし、実際の結果を取得
            // ここではList<Recipient>にキャストできない（理由は？）
            var objArray = new object[] { testMeeting, Utility.OutlookItemType.Mail, false };
            object actualObj = mi2.Invoke(obj2, objArray);

            // テスト対象メソッドの返り値をList<Recipient>型にする
            List<Recipient> actualRecList = new List<Recipient>();
            IEnumerable<Recipient> actualEnumList = (IEnumerable<Recipient>)actualObj;

            foreach (var actual in actualEnumList)
            {
                actualRecList.Add(actual);
            }

            // 期待結果を入れるリスト
            List<Recipient> expectedRecList = new List<Recipient>();

            // 期待結果1のデータをリストに追加
            Recipient expectedRec1 = Substitute.For<Recipient>();
            expectedRec1.Address.Returns("testemailaddress1@example.com");
            expectedRec1.Sendable.Returns(true);
            expectedRec1.Type.Returns((int)OlMailRecipientType.olTo);
            expectedRecList.Add(expectedRec1);

            // 期待結果2のデータをリストに追加
            Recipient expectedRec2 = Substitute.For<Recipient>();
            expectedRec2.Address.Returns("testemailaddress2@example.com");
            expectedRec2.Sendable.Returns(true);
            expectedRec2.Type.Returns((int)OlMailRecipientType.olCC);
            expectedRecList.Add(expectedRec2);

            // actualとexpectedのリストを比較
            Assert.AreEqual(actualRecList.Count, expectedRecList.Count);

            Assert.That(actualRecList[0].Address, Is.EqualTo(expectedRecList[0].Address));
            Assert.That(actualRecList[0].Sendable, Is.EqualTo(expectedRecList[0].Sendable));
            Assert.That(actualRecList[0].Type, Is.EqualTo(expectedRecList[0].Type));

            Assert.That(actualRecList[1].Address, Is.EqualTo(expectedRecList[1].Address));
            Assert.That(actualRecList[1].Sendable, Is.EqualTo(expectedRecList[1].Sendable));
            Assert.That(actualRecList[1].Type, Is.EqualTo(expectedRecList[1].Type));

            // ref引数のtypeが正しいことを確認
            Assert.That(objArray[1], Is.EqualTo(Utility.OutlookItemType.Meeting));
        }

        /// <summary>
        /// 
        /// </summary>
        [Test]
        public void GetRecipientsTest3()
        {

        }

        /// <summary>
        /// 
        /// </summary>
        [Test]
        public void GetRecipientsTest4()
        {

        }

        /// <summary>
        /// 
        /// </summary>
        [Test]
        public void GetRecipientsTest5()
        {

        }

        /// <summary>
        /// 
        /// </summary>
        [Test]
        public void GetRecipientsTest6()
        {

        }

        /// <summary>
        /// 
        /// </summary>
        [Test]
        public void GetRecipientsTest7()
        {

        }



        class TestAddIn : OutlookAddIn
        {
            public BindingContext BindingContext
            { get; set; }

            public ControlBindingsCollection DataBindings
            { get; }

            public ICachedDataProvider DataHost
            { get; }

            public IAddInExtension DefaultExtension
            {
                get
                {
                    return Substitute.For<IAddInExtension>();
                }
            }

            public IAddInExtension Extension
            {
                get
                {
                    return Substitute.For<IAddInExtension>();
                }
            }

            public IServiceProvider HostContext
            { get; }

            public IHostItemProvider ItemProvider
            { get; }

            public ISite Site
            { get; set; }

            public dynamic Tag
            { get; set; }

            public event EventHandler BindingContextChanged;
            public event EventHandler Disposed;
            public event FormRegionFactoryResolveEventHandler FormRegionFactoryResolve;
            public event EventHandler Shutdown;
            public event EventHandler Startup;

            public void Dispose()
            {

            }

            public IList<IFormRegion> GetFormRegions()
            {
                throw new NotImplementedException();
            }

            public IList<IFormRegion> GetFormRegions(Inspector inspector, Type customCollectionType)
            {
                throw new NotImplementedException();
            }

            public IList<IFormRegion> GetFormRegions(Explorer explorer, Type customCollectionType)
            {
                throw new NotImplementedException();
            }
        }

        class TestFactory : Microsoft.Office.Tools.Outlook.Factory
        {
            public AddIn CreateAddIn(IServiceProvider serviceProvider, IHostItemProvider hostItemProvider, string primaryCookie, string identifier, object containerComponent, IAddInExtension extension)
            {
                return new TestAddIn();
            }

            public CustomTaskPaneCollection CreateCustomTaskPaneCollection(IServiceProvider serviceProvider, IHostItemProvider hostItemProvider, string primaryCookie, string identifier, object containerComponent)
            {
                throw new NotImplementedException();
            }

            public IList<IFormRegion> CreateFormRegionCollection()
            {
                throw new NotImplementedException();
            }

            public FormRegionControl CreateFormRegionControl(FormRegion region, IExtension extension)
            {
                throw new NotImplementedException();
            }

            public FormRegionCustomAction CreateFormRegionCustomAction()
            {
                throw new NotImplementedException();
            }

            public FormRegionCustomAction CreateFormRegionCustomAction(string name)
            {
                throw new NotImplementedException();
            }

            public FormRegionInitializingEventArgs CreateFormRegionInitializingEventArgs(object outlookItem, OlFormRegionMode formRegionMode, OlFormRegionSize formRegionSize, bool cancel)
            {
                throw new NotImplementedException();
            }

            public FormRegionManifest CreateFormRegionManifest()
            {
                throw new NotImplementedException();
            }

            public ImportedFormRegion CreateImportedFormRegion(FormRegion region, IImportedFormRegionExtension extension)
            {
                throw new NotImplementedException();
            }

            public SmartTagCollection CreateSmartTagCollection(IServiceProvider serviceProvider, IHostItemProvider hostItemProvider, string primaryCookie, string identifier, object containerComponent)
            {
                throw new NotImplementedException();
            }

            public RibbonFactory GetRibbonFactory()
            {
                throw new NotImplementedException();
            }
        }

        public abstract class TestApplication : Microsoft.Office.Interop.Outlook.Application
        {
            public abstract Microsoft.Office.Core.AnswerWizard AnswerWizard { get; }
            public abstract Microsoft.Office.Interop.Outlook.Application Application { get; }
            public abstract Microsoft.Office.Core.IAssistance Assistance { get; }
            public abstract Microsoft.Office.Core.Assistant Assistant { get; }
            public abstract OlObjectClass Class { get; }
            public abstract Microsoft.Office.Core.COMAddIns COMAddIns { get; }
            public abstract string DefaultProfileName { get; }
            public abstract Explorers Explorers { get; }
            public abstract Microsoft.Office.Core.MsoFeatureInstall FeatureInstall { get; set; }
            public abstract Inspectors Inspectors { get; }
            public abstract bool IsTrusted { get; }
            public abstract Microsoft.Office.Core.LanguageSettings LanguageSettings { get; }
            public abstract string Name { get; }
            public abstract dynamic Parent { get; }
            public abstract Microsoft.Office.Core.PickerDialog PickerDialog { get; }
            public abstract string ProductCode { get; }
            public abstract Reminders Reminders { get; }
            public abstract NameSpace Session { get; }
            public abstract TimeZones TimeZones { get; }
            public abstract string Version { get; }

            public abstract event ApplicationEvents_11_AdvancedSearchCompleteEventHandler AdvancedSearchComplete;
            public abstract event ApplicationEvents_11_AdvancedSearchStoppedEventHandler AdvancedSearchStopped;
            public abstract event ApplicationEvents_11_AttachmentContextMenuDisplayEventHandler AttachmentContextMenuDisplay;
            public abstract event ApplicationEvents_11_BeforeFolderSharingDialogEventHandler BeforeFolderSharingDialog;
            public abstract event ApplicationEvents_11_ContextMenuCloseEventHandler ContextMenuClose;
            public abstract event ApplicationEvents_11_FolderContextMenuDisplayEventHandler FolderContextMenuDisplay;
            public abstract event ApplicationEvents_11_ItemContextMenuDisplayEventHandler ItemContextMenuDisplay;
            public abstract event ApplicationEvents_11_ItemLoadEventHandler ItemLoad;
            public abstract event ApplicationEvents_11_ItemSendEventHandler ItemSend;
            public abstract event ApplicationEvents_11_MAPILogonCompleteEventHandler MAPILogonComplete;
            public abstract event ApplicationEvents_11_NewMailEventHandler NewMail;
            public abstract event ApplicationEvents_11_NewMailExEventHandler NewMailEx;
            public abstract event ApplicationEvents_11_OptionsPagesAddEventHandler OptionsPagesAdd;
            public abstract event ApplicationEvents_11_ReminderEventHandler Reminder;
            public abstract event ApplicationEvents_11_ShortcutContextMenuDisplayEventHandler ShortcutContextMenuDisplay;
            public abstract event ApplicationEvents_11_StartupEventHandler Startup;
            public abstract event ApplicationEvents_11_StoreContextMenuDisplayEventHandler StoreContextMenuDisplay;
            public abstract event ApplicationEvents_11_ViewContextMenuDisplayEventHandler ViewContextMenuDisplay;

            event ApplicationEvents_11_QuitEventHandler ApplicationEvents_11_Event.Quit
            {
                add
                {
                    throw new NotImplementedException();
                }

                remove
                {
                    throw new NotImplementedException();
                }
            }

            public abstract Explorer ActiveExplorer();
            public abstract Inspector ActiveInspector();
            public abstract dynamic ActiveWindow();
            public abstract Search AdvancedSearch(string Scope, object Filter, object SearchSubFolders, object Tag);
            public abstract dynamic CopyFile(string FilePath, string DestFolderPath);

            /// <summary>
            /// テスト対象のクラス(e.g. Office365Contact)が実行するクラス
            /// </summary>
            /// <param name="ItemType">生成するアイテムのタイプ</param>
            /// <returns></returns>
            public dynamic CreateItem(OlItemType ItemType)
            {
                return CreateItemHon(ItemType);
            }

            /// <summary>
            /// dynamicを返さないために作ったメソッド
            /// virtualをつけることで、Substituteに制御を渡す
            /// </summary>
            /// <param name="ItemType">生成するアイテムのタイプ</param>
            /// <returns></returns>
            public virtual ContactItem CreateItemHon(OlItemType ItemType)
            {
                return null;
            }

            public abstract dynamic CreateItemFromTemplate(string TemplatePath, object InFolder);
            public abstract dynamic CreateObject(string ObjectName);
            public abstract NameSpace GetNamespace(string Type);
            public abstract void GetNewNickNames(ref object pvar);
            public abstract dynamic GetObjectReference(object Item, OlReferenceType ReferenceType);
            public abstract bool IsSearchSynchronous(string LookInFolders);
            public abstract void Quit();
            public abstract void RefreshFormRegionDefinition(string RegionName);
        }


    }

}
