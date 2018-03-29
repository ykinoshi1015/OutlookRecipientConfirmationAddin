using System;
using NUnit.Framework;
using NSubstitute;
using Microsoft.Office.Interop.Outlook;
using OutlookRecipientConfirmationAddin;
using System.Reflection;

namespace ORCAUnitTest
{

    /// <summary>
    /// Utilityクラス GetSenderInformationメソッドのテストクラス
    /// </summary>
    /// <remarks>
    /// アイテムから、送信者情報を取得するメソッドの単体テストコード
    /// </remarks>
    [TestFixture]
    public class GetSenderInformationUnitTest
    {
        // テスト対象のクラスのインスタンス
        private object obj;

        // テスト対象のメソッド属性
        private MethodInfo mi;

        // テストするアイテムのモック
        private MailItem testMail;
        private MeetingItem testMeeting;
        private AppointmentItem testAppointment;
        private SharingItem testSharing;
        private TestReportItem testReport;
        private DocumentItem testDocument;

        // テスト対象のクラスで使われる変数のモック
        private Recipient testRec;
        private AddressEntry testAdd;
        private ExchangeUser testExchUser;
        private Application testApp;
        private NameSpace testNs;

        /// <summary>
        /// テスト時に一度だけ実行される処理
        /// </summary>
        /// <remarks>
        /// アセンブリの読み込み、Typeの取得、モックの作成など
        /// </remarks>
        [OneTimeSetUp]
        public void Init()
        {
            // テスト対象のクラス内で使われる変数のモック
            testRec = Substitute.For<Recipient>();
            //testAdd = Substitute.For<AddressEntry>();
            testExchUser = Substitute.For<ExchangeUser>();

            // テスト用のXXXItemを、モックで作成
            testMail = Substitute.For<MailItem>();
            testMeeting = Substitute.For<MeetingItem>();
            testAppointment = Substitute.For<AppointmentItem>();
            testSharing = Substitute.For<SharingItem>();
            testReport = Substitute.For<TestReportItem>();
            testDocument = Substitute.For<DocumentItem>();

            // ------------------------------------------------------------------------------------------------
            // VSで実行する場合
            // ------------------------------------------------------------------------------------------------

            //// Factoryクラスは、自作のTestFactoryクラス＆その中で使うTestAddInクラスがないとうまくいかない
            //// ThisAddInクラスのタイプを取得
            //TestFactory testFactory = new TestFactory();
            //IServiceProvider testService = Substitute.For<IServiceProvider>();
            //ThisAddIn testAddIn = new ThisAddIn(testFactory, testService);

            //// リフレクション
            //// アセンブリを読み込み、モジュールを取得
            //Assembly asm = Assembly.LoadFrom(@".\ORCAUnitTest\bin\Debug\OutlookRecipientConfirmationAddin.dll");

            // ------------------------------------------------------------------------------------------------
            // batで、このプロジェクトのテストをまとめて実行する場合
            // ------------------------------------------------------------------------------------------------

            // 共通で使うThisAddInクラスを取得
            ThisAddIn testAddIn = GetContactItemUnitTest.testAddIn;

            // リフレクション
            // アセンブリを読み込み、モジュールを取得
            Assembly asm = Assembly.LoadFrom(@".\OutlookRecipientConfirmationAddin.dll");

            // ------------------------------------------------------------------------------------------------

            Module mod = asm.GetModule("OutlookRecipientConfirmationAddin.dll");
            Type typeThisAddIn = testAddIn.GetType();

            // Applicaitionのモック作成
            FieldInfo fieldApp = typeThisAddIn.GetField("Application", BindingFlags.NonPublic | BindingFlags.Instance);
            testApp = Substitute.For<TestApplication>();
            fieldApp.SetValue(testAddIn, testApp);

            // Sessionのモック作成
            testNs = Substitute.For<NameSpace>();
            testNs.CreateRecipient(Arg.Any<string>()).Returns(testRec);
            testApp.Session.Returns(testNs);

            // ------------------------------------------------------------------------------------------------
            // このクラスを単独で実行する場合
            // ------------------------------------------------------------------------------------------------

            //// Globalsのタイプと、ThisAddInプロパティを取得
            //Type typeGlobal = mod.GetType("OutlookRecipientConfirmationAddin.Globals");
            //PropertyInfo testProp = typeGlobal.GetProperty("ThisAddIn", BindingFlags.NonPublic | BindingFlags.Static);
            //// ThisAddinプロパティに、モックなどを使って作った値をセットする
            //testProp.SetValue(null, testAddIn);

            // ------------------------------------------------------------------------------------------------

            // テスト対象のクラス（Utility）のタイプを取得
            Type type = mod.GetType("OutlookRecipientConfirmationAddin.Utility");
            // インスタンスを生成し、メソッドにアクセスできるようにする
            obj = Activator.CreateInstance(type);
            mi = type.GetMethod("GetSenderInfomation");
        }


        /// <summary>
        /// <para>アイテムが、MailItemの場合</para>
        /// <para>（Senderプロパティが取得できる）</para>
        /// <para>（Senderプロパティから、ExchangeUserが取得できる）</para>
        /// </summary>
        /// <remarks>
        /// 【期待結果】
        /// <para>senderInformationDtoのfullName, division, companyName, recipientTypeが取得できる</para>
        /// <para>jobTitleが""になる</para>
        /// </remarks>
        [Test]
        public void GetSenderInfoMailTest1()
        {
            string testName = "Kosaka Kenta (小坂 健太)";
            string testCompanyName = "リコーITソリューションズ";
            string testDepartment = "ビジネスソリューションズ事業部 システム開発センター 第１開発部 第１グループ";
            string testJobTitle = "担当";
            string expectedJobTitle = "";

            testAdd = Substitute.For<AddressEntry>();

            // モックでつかうデータを用意
            testMail.Sender = testAdd;
            testAdd.Address = "kenta.kosaka@jp.ricoh.com";

            // モックのReturn値を設定
            testRec.AddressEntry.Returns(testAdd);

            testExchUser.Name.Returns(testName);
            testExchUser.Department.Returns(testDepartment);
            testExchUser.CompanyName.Returns(testCompanyName);
            testExchUser.JobTitle.Returns(testJobTitle);
            testAdd.GetExchangeUser().Returns(testExchUser);

            // テストするメソッドにアクセスし、実際の結果を取得
            RecipientInformationDto actual = (RecipientInformationDto)mi.Invoke(obj, new object[] { testMail });

            // 期待結果
            RecipientInformationDto expected = new RecipientInformationDto(testName, testDepartment, testCompanyName, expectedJobTitle, OlMailRecipientType.olOriginator);

            // actualとexpectedを比較
            CompareRecInfoDto(actual, expected);
        }

        /// <summary>
        /// <para>アイテムが、MailItemの場合</para>
        /// <para>（Senderプロパティが取得できる）</para>
        /// <para>（Senderプロパティから、ExchangeUserが取得できない）</para>
        /// </summary>
        /// <remarks>
        /// 【期待結果】
        /// <para>senderInformationDtoのrecipientTypeとemailAddressが取得できる</para>
        /// </remarks>
        [Test]
        public void GetSenderInfoMailTest2()
        {
            string testName = "Kosaka Kenta (小坂 健太)";
            string testEmailAddress = "kenta.kosaka@jp.ricoh.com";
            string expectedNameAndAddress = string.Format("{0}<{1}>", testName, testEmailAddress);

            testAdd = Substitute.For<AddressEntry>();

            // モックでつかうデータを用意
            testMail.Sender = testAdd;
            testAdd.Address = "kenta.kosaka@jp.ricoh.com";

            // モックのReturn値を設定
            testRec.AddressEntry.Returns(testAdd);

            // getExchangeUserメソッドで、ExchangeUserが見つからない
            testAdd.GetExchangeUser().Returns(x => { throw new System.Exception(); });

            testMail.SenderName.Returns(testName);
            testMail.SenderEmailAddress.Returns(testEmailAddress);

            // テストするメソッドにアクセスし、実際の結果を取得
            RecipientInformationDto actual = (RecipientInformationDto)mi.Invoke(obj, new object[] { testMail });


            // 期待結果
            RecipientInformationDto expected = new RecipientInformationDto(expectedNameAndAddress, OlMailRecipientType.olOriginator);

            // actualとexpectedを比較
            CompareRecInfoDto(actual, expected);
        }

        /// <summary>
        /// <para>アイテムが、MailItemの場合</para>
        /// <para>（Senderプロパティがnull）</para>
        /// <para>（SenderEamilAddressプロパティから、ExchangeUserが取得できる）</para>
        /// </summary>
        /// <remarks>
        /// 【期待結果】
        /// <para>senderInformationDtoのfullName, division, companyName, recipientTypeが取得できる</para>
        /// <para>jobTitleが"部長"になる</para>
        /// </remarks>
        [Test]
        public void GetSenderInfoMailTest3()
        {
            string testName = "Kobayashi Gen (小林 元)";
            string testCompanyName = "リコーITソリューションズ";
            string testDepartment = "ビジネスソリューションズ事業部 システム開発センター 第１開発部 第１グループ";
            string testJobTitle = "部長";

            testAdd = Substitute.For<AddressEntry>();

            // モックでつかうデータを用意
            testMail.Sender = null;
            testMail.SenderEmailAddress.Returns("gen.kobayashi@jp.ricoh.com");

            testAdd = Substitute.For<AddressEntry>();

            // モックのReturn値を設定
            testRec.AddressEntry.Returns(testAdd);

            testExchUser.Name.Returns(testName);
            testExchUser.Department.Returns(testDepartment);
            testExchUser.CompanyName.Returns(testCompanyName);
            testExchUser.JobTitle.Returns(testJobTitle);
            testAdd.GetExchangeUser().Returns(testExchUser);

            // テストするメソッドにアクセスし、実際の結果を取得
            RecipientInformationDto actual = (RecipientInformationDto)mi.Invoke(obj, new object[] { testMail });

            // 期待結果
            RecipientInformationDto expected = new RecipientInformationDto(testName, testDepartment, testCompanyName, testJobTitle, OlMailRecipientType.olOriginator);

            // actualとexpectedを比較
            CompareRecInfoDto(actual, expected);
        }

        /// <summary>
        /// <para>アイテムが、MailItemの場合</para>
        /// <para>（Senderプロパティがnull）</para>
        /// <para>（SenderEamilAddressプロパティはnullでないが、ExchangeUserが取得できない）</para>
        /// </summary>
        /// <remarks>
        /// 【期待結果】
        /// <para>senderInformationDtoのrecipientTypeとemailAddressが取得できる</para>
        /// </remarks>
        [Test]
        public void GetSenderInfoMailTest4()
        {
            string testName = "Kosaka Kenta (小坂 健太)";
            string testEmailAddress = "kenta.kosaka@jp.ricoh.com";
            string expectedNameAndAddress = string.Format("{0}<{1}>", testName, testEmailAddress);

            testAdd = Substitute.For<AddressEntry>();

            // モックでつかうデータを用意
            testMail.Sender = null;
            testMail.SenderEmailAddress.Returns(testEmailAddress);
            testMail.SenderName.Returns(testName);

            // モックのReturn値を設定
            testRec.AddressEntry.Returns(testAdd);

            // getExchangeUserメソッドで、ExchangeUserが見つからない
            testAdd.GetExchangeUser().Returns(x => { throw new System.Exception(); });

            // テストするメソッドにアクセスし、実際の結果を取得
            RecipientInformationDto actual = (RecipientInformationDto)mi.Invoke(obj, new object[] { testMail });

            // 期待結果
            RecipientInformationDto expected = new RecipientInformationDto(expectedNameAndAddress, OlMailRecipientType.olOriginator);

            // actualとexpectedを比較
            CompareRecInfoDto(actual, expected);

        }

        /// <summary>
        /// <para>アイテムが、MailItemの場合</para>
        /// <para>（Senderプロパティ/SenderEamilAddressプロパティから、ExchangeUserが取得できない）</para>
        /// </summary>
        /// <remarks>
        /// 【期待結果】
        /// <para>senderInformationDtoがnull</para>
        /// </remarks>
        [Test]
        public void GetSenderInfoMailTest5()
        {
            testAdd = Substitute.For<AddressEntry>();

            // モックでつかうデータを用意
            testMail.Sender = null;
            testMail.SenderEmailAddress.Returns((string)null);
            testMail.SenderName.Returns((string)null);

            // モックのReturn値を設定
            testRec.AddressEntry.Returns(testAdd);

            // テストするメソッドにアクセスし、実際の結果を取得
            RecipientInformationDto actual = (RecipientInformationDto)mi.Invoke(obj, new object[] { testMail });

            // メソッドの戻り値がnullであることを確認
            Assert.IsNull(actual);
        }

        /// <summary>
        /// <para>アイテムが、MeetingItemの場合</para>
        /// <para>（送信者のAddressEntryが取得できる）</para>
        /// <para>（送信者のAddressEntryから、ExchangeUserも取得できる）</para>
        /// </summary>
        /// <remarks>
        /// 【期待結果】
        /// <para> senderInformationDtoのfullName, division, companyName, recipientTypeが取得できる</para>
        /// <para> jobTitleが""になる</para>
        /// </remarks>
        [Test]
        public void GetSenderInfoMeetingTest1()
        {
            string testName = "Kosaka Kenta (小坂 健太)";
            string testCompanyName = "リコーITソリューションズ";
            string testDepartment = "ビジネスソリューションズ事業部 システム開発センター 第１開発部 第１グループ";
            string testJobTitle = "担当";
            string testEmailAddress = "kenta.kosaka@jp.ricoh.com";
            string expectedJobTitle = "";

            testAdd = Substitute.For<AddressEntry>();

            // モックのReturn値を設定
            testMeeting.SenderEmailAddress.Returns(testEmailAddress);
            testRec.AddressEntry.Returns(testAdd);

            testAdd.GetExchangeUser().Returns(testExchUser);
            testExchUser.Name.Returns(testName);
            testExchUser.Department.Returns(testDepartment);
            testExchUser.CompanyName.Returns(testCompanyName);
            testExchUser.JobTitle.Returns(testJobTitle);

            // テストするメソッドにアクセスし、実際の結果を取得
            RecipientInformationDto actual = (RecipientInformationDto)mi.Invoke(obj, new object[] { testMeeting });

            // 期待結果
            RecipientInformationDto expected = new RecipientInformationDto(testName, testDepartment, testCompanyName, expectedJobTitle, OlMailRecipientType.olOriginator);

            // actualとexpectedを比較
            CompareRecInfoDto(actual, expected);
        }

        /// <summary>
        /// <para>アイテムが、MeetingItemの場合</para>
        /// <para>（送信者のAddressEntryが取得できる）</para>
        /// <para>（送信者のAddressEntryからExchangeUserが取得できない(例外が発生)）</para>
        /// </summary>
        /// <remarks>
        /// 【期待結果】
        /// <para> senderInformationDtoのrecipientTypeとemailAddressが取得できる</para>
        /// </remarks>
        [Test]
        public void GetSenderInfoMeetingTest2()
        {
            string testName = "Kosaka Kenta (小坂 健太)";
            string testEmailAddress = "kenta.kosaka@jp.ricoh.com";
            string expectedNameAndAddress = string.Format("{0}<{1}>", testName, testEmailAddress);

            testAdd = Substitute.For<AddressEntry>();

            // モックでつかうデータを用意
            testMeeting.SenderName.Returns(testName);
            testMeeting.SenderEmailAddress.Returns(testEmailAddress);

            // モックのReturn値を設定
            testRec.AddressEntry.Returns(testAdd);

            // GetExchangeUserメソッドで、例外が発生
            testAdd.GetExchangeUser().Returns(x => { throw new System.Exception(); });

            // テストするメソッドにアクセスし、実際の結果を取得
            RecipientInformationDto actual = (RecipientInformationDto)mi.Invoke(obj, new object[] { testMeeting });

            // 期待結果
            RecipientInformationDto expected = new RecipientInformationDto(expectedNameAndAddress, OlMailRecipientType.olOriginator);

            // actualとexpectedを比較
            CompareRecInfoDto(actual, expected);
        }

        /// <summary>
        /// <para>アイテムが、MeetingItemの場合</para>
        /// <para>（送信者のAddressEntryが取得できない）</para>
        /// </summary>
        /// <remarks>
        /// 【期待結果】
        /// <para>senderInformationDtoがnull</para>
        /// </remarks>
        [Test]
        public void GetSenderInfoMeetingTest3()
        {
            string testEmailAddress = "kenta.kosaka@jp.ricoh.com";

            testAdd = Substitute.For<AddressEntry>();

            // モックでつかうデータを用意
            testMeeting.SenderName.Returns((string)null);
            testMeeting.SenderEmailAddress.Returns(testEmailAddress);

            // モックのReturn値を設定
            testRec.AddressEntry.Returns(testAdd);

            // GetExchangeUserメソッドで、例外が発生
            testAdd.GetExchangeUser().Returns(x => { throw new System.Exception(); });

            // テストするメソッドにアクセスし、実際の結果を取得
            RecipientInformationDto actual = (RecipientInformationDto)mi.Invoke(obj, new object[] { testMeeting });

            // メソッドの戻り値がnullであることを確認
            Assert.IsNull(actual);
        }

        /// <summary>
        /// <para>アイテムが、MeetingItemの場合</para>
        /// <para>（送信者のAddressEntryが取得できる）</para>
        /// <para>（送信者のAddressEntryからExchangeUserが取得できない(ExchangeUserがnull)）</para>
        /// <para>（RecipientのNameプロパティですでに「名前(メールアドレス)」の形式になっている）</para>
        /// </summary>
        /// <remarks>
        /// 【期待結果】
        /// <para>senderInformationDtoのrecipientTypeとemailAddressが取得できる</para>
        /// </remarks>
        [Test]
        public void GetSenderInfoMeetingTest4()
        {
            string testName = "Kosaka Kenta (小坂 健太)";
            string testEmailAddress = "kenta.kosaka@jp.ricoh.com";
            string testNameAndAddress = "小坂 健太 (kenta.kosaka@jp.ricoh.com)";

            testAdd = Substitute.For<AddressEntry>();

            // モックでつかうデータを用意
            testMeeting.SenderName.Returns(testName);
            testMeeting.SenderEmailAddress.Returns(testEmailAddress);

            // モックのReturn値を設定
            testRec.AddressEntry.Returns(testAdd);
            testRec.Name.Returns(testNameAndAddress);

            // GetExchangeUserメソッドで、例外が発生
            testAdd.GetExchangeUser().Returns((ExchangeUser)null);

            // テストするメソッドにアクセスし、実際の結果を取得
            RecipientInformationDto actual = (RecipientInformationDto)mi.Invoke(obj, new object[] { testMeeting });

            // 期待結果
            RecipientInformationDto expected = new RecipientInformationDto(testNameAndAddress, OlMailRecipientType.olOriginator);

            // actualとexpectedを比較
            CompareRecInfoDto(actual, expected);
        }

        /// <summary>
        /// <para>アイテムが、MeetingItemの場合</para>
        /// <para>（送信者のAddressEntryが取得できる）</para>
        /// <para>（送信者のAddressEntryからExchangeUserが取得できない(ExchangeUserがnull)）</para>
        /// <para>（表示用に"名前＜メールアドレス＞"の形式の文字列にする）</para>
        /// </summary>
        /// <remarks>
        /// 【期待結果】
        /// <para>senderInformationDtoのrecipientTypeとemailAddressが取得できる</para>
        /// </remarks>
        [Test]
        public void GetSenderInfoMeetingTest5()
        {
            string testAddress = "yasuyuki.kinoshita@jrits.ricoh.co.jp";
            string testName = "Yasuyuki Kinoshita / R / RSI";
            string expectedNameAndAddress = string.Format("{0}<{1}>", testName, testAddress);

            testAdd = Substitute.For<AddressEntry>();

            // モックでつかうデータを用意
            testMeeting.SenderName.Returns(testName);
            testMeeting.SenderEmailAddress.Returns(testAddress);

            // モックのReturn値を設定
            testRec.AddressEntry.Returns(testAdd);
            testRec.Name.Returns(testName);
            testRec.Address.Returns(testAddress);

            // GetExchangeUserメソッドで、例外が発生
            testAdd.GetExchangeUser().Returns((ExchangeUser)null);

            // テストするメソッドにアクセスし、実際の結果を取得
            RecipientInformationDto actual = (RecipientInformationDto)mi.Invoke(obj, new object[] { testMeeting });

            // 期待結果
            RecipientInformationDto expected = new RecipientInformationDto(expectedNameAndAddress, OlMailRecipientType.olOriginator);

            // actualとexpectedを比較
            CompareRecInfoDto(actual, expected);
        }
        
        /// <summary>
        /// <para>アイテムが、AppointmentItemの場合</para>
        /// <para>（Recipients[1]のExchangeUserが取得できる）</para>
        /// </summary>
        /// <remarks>
        /// 【期待結果】
        /// <para>senderInformationDtoのfullName, division, companyName, recipientTypeが取得できる</para>
        /// <para>jobTitleが""になる</para>
        /// </remarks>
        [Test]
        public void GetSenderInfoAppointTest1()
        {
            string testName = "Kosaka Kenta (小坂 健太)";
            string testCompanyName = "リコーITソリューションズ";
            string testDepartment = "ビジネスソリューションズ事業部 システム開発センター 第１開発部 第１グループ";
            string testJobTitle = "担当";
            string expectedJobTitle = "";

            testAdd = Substitute.For<AddressEntry>();

            // モックのReturn値を設定
            testAppointment.Recipients[1].AddressEntry.Returns(testAdd);

            testExchUser.Name.Returns(testName);
            testExchUser.Department.Returns(testDepartment);
            testExchUser.CompanyName.Returns(testCompanyName);
            testExchUser.JobTitle.Returns(testJobTitle);
            testAdd.GetExchangeUser().Returns(testExchUser);

            // テストするメソッドにアクセスし、実際の結果を取得
            RecipientInformationDto actual = (RecipientInformationDto)mi.Invoke(obj, new object[] { testAppointment });

            // 期待結果
            RecipientInformationDto expected = new RecipientInformationDto(testName, testDepartment, testCompanyName, expectedJobTitle, OlMailRecipientType.olOriginator);

            // actualとexpectedを比較
            CompareRecInfoDto(actual, expected);
        }

        /// <summary>
        /// <para>アイテムが、AppointmentItemの場合</para>
        /// <para>（Recipients[1]のExchangeUserが取得できない）</para>
        /// <para>（現在のユーザの AddressEntryから、ExchangeUserが取得できる）</para>
        /// </summary>
        /// <remarks>
        /// 【期待結果】
        /// <para>senderInformationDtoのfullName, division, companyName, recipientTypeが取得できる</para>
        /// <para>jobTitleが""になる</para>
        /// </remarks>
        [Test]
        public void GetSenderInfoAppointTest2()
        {
            string testName = "Kosaka Kenta (小坂 健太)";
            string testCompanyName = "リコーITソリューションズ";
            string testDepartment = "ビジネスソリューションズ事業部 システム開発センター 第１開発部 第１グループ";
            string testJobTitle = "担当";
            string expectedJobTitle = "";

            testAdd = Substitute.For<AddressEntry>();
            testAppointment = Substitute.For<AppointmentItem>();

            // Recipients[1]のExchangeUserで例外が発生
            testAppointment.Recipients[1].AddressEntry.Returns(x => { throw new System.Exception(); });

            // モックのReturn値を設定
            testNs.CurrentUser.AddressEntry.Returns(testAdd);

            testExchUser.Name.Returns(testName);
            testExchUser.Department.Returns(testDepartment);
            testExchUser.CompanyName.Returns(testCompanyName);
            testExchUser.JobTitle.Returns(testJobTitle);
            testAdd.GetExchangeUser().Returns(testExchUser);

            // テストするメソッドにアクセスし、実際の結果を取得
            RecipientInformationDto actual = (RecipientInformationDto)mi.Invoke(obj, new object[] { testAppointment });

            // 期待結果
            RecipientInformationDto expected = new RecipientInformationDto(testName, testDepartment, testCompanyName, expectedJobTitle, OlMailRecipientType.olOriginator);

            // actualとexpectedを比較
            CompareRecInfoDto(actual, expected);
        }

        /// <summary>
        /// <para>アイテムが、AppointmentItemの場合</para>
        /// <para>（Recipients[1]のExchangeUserが取得できない）</para>
        /// <para>（現在のユーザの AddressEntryから、ExchangeUserが取得できない）</para>
        /// </summary>
        /// <remarks>
        /// 【期待結果】
        /// <para>senderInformationDtoのrecipientTypeとemailAddressが取得できる</para>
        /// </remarks>
        [Test]
        public void GetSenderInfoAppointTest3()
        {
            string testName = "Kinoshita Yasuyuki (木下 康行)";
            string testEmailAddress = "yasuyuki.kinoshita@jp.ricoh.com";
            string expectedNameAndAddress = string.Format("{0}<{1}>", testName, testEmailAddress);

            testAdd = Substitute.For<AddressEntry>();
            testAppointment = Substitute.For<AppointmentItem>();

            // Recipients[1]のExchangeUserで例外が発生
            testAppointment.Recipients[1].AddressEntry.Returns(x => { throw new System.Exception(); });

            // モックのReturn値を設定
            testNs.CurrentUser.AddressEntry.Returns(testAdd);
            testAdd.Name = testName;
            testAdd.Address = testEmailAddress;
            testAdd.GetExchangeUser().Returns((ExchangeUser)null);

            // テストするメソッドにアクセスし、実際の結果を取得
            RecipientInformationDto actual = (RecipientInformationDto)mi.Invoke(obj, new object[] { testAppointment });

            // 期待結果
            RecipientInformationDto expected = new RecipientInformationDto(expectedNameAndAddress, OlMailRecipientType.olOriginator);

            // actualとexpectedを比較
            CompareRecInfoDto(actual, expected);
        }
       
        /// <summary>
        /// <para>アイテムが、AppointmentItemの場合</para>
        /// <para>（Recipients[1]のExchangeUserが取得できない）</para>
        /// <para>(現在のユーザの AddressEntryが取得できない)</para>
        /// </summary>
        /// <remarks>
        /// 【期待結果】
        /// <para>senderInformationDtoがnull</para>
        /// </remarks>
        [Test]
        public void GetSenderInfoAppointTest4()
        {

            testAdd = Substitute.For<AddressEntry>();
            testAppointment = Substitute.For<AppointmentItem>();

            // Recipients[1]のExchangeUserで例外が発生
            testAppointment.Recipients[1].AddressEntry.Returns(x => { throw new System.Exception(); });

            // モックのReturn値を設定
            testNs.CurrentUser.AddressEntry.Returns((AddressEntry)null);

            // テストするメソッドにアクセスし、実際の結果を取得
            RecipientInformationDto actual = (RecipientInformationDto)mi.Invoke(obj, new object[] { testAppointment });

            // メソッドの戻り値がnullであることを確認
            Assert.IsNull(actual);
        }

        /// <summary>
        /// <para>アイテムが、ReportItemの場合</para>
        /// <para>（Recipients[1]のExchangeUserが取得できない）</para>
        /// <para>（現在のユーザの AddressEntryから、ExchangeUserが取得できない）</para>
        /// </summary>
        /// <remarks>
        /// 【期待結果】
        /// <para>senderInformationDtoがnull</para>
        /// </remarks>
        [Test]
        public void GetSenderInfoReportTest1()
        {
            string testSender = "Microsoft Outlook";

            testAdd = Substitute.For<AddressEntry>();

            // テストするメソッドにアクセスし、実際の結果を取得
            RecipientInformationDto actual = (RecipientInformationDto)mi.Invoke(obj, new object[] { testReport });

            // 期待結果
            RecipientInformationDto expected = new RecipientInformationDto(testSender, OlMailRecipientType.olOriginator);

            // actualとexpectedを比較
            CompareRecInfoDto(actual, expected);
        }
        
        /// <summary>
        /// <para>アイテムが、SharingItemの場合</para>
        /// <para>（送信者のAddressEntryが取得できる）</para>
        /// <para>（送信者のAddressEntryから、ExchangeUserも取得できる場合）</para>
        /// </summary>
        /// <remarks>
        /// 【期待結果】
        /// <para>senderInformationDtoのfullName, division, companyName, recipientTypeが取得できる</para>
        /// <para>jobTitleが""になる</para>
        /// </remarks>
        [Test]
        public void GetSenderInfoSharingTest1()
        {
            string testName = "Kosaka Kenta (小坂 健太)";
            string testCompanyName = "リコーITソリューションズ";
            string testDepartment = "ビジネスソリューションズ事業部 システム開発センター 第１開発部 第１グループ";
            string testJobTitle = "担当";
            string testEmailAddress = "kenta.kosaka@jp.ricoh.com";
            string expectedJobTitle = "";

            testAdd = Substitute.For<AddressEntry>();

            // モックのReturn値を設定
            testSharing.SenderEmailAddress.Returns(testEmailAddress);
            testRec.AddressEntry.Returns(testAdd);

            testAdd.GetExchangeUser().Returns(testExchUser);
            testExchUser.Name.Returns(testName);
            testExchUser.Department.Returns(testDepartment);
            testExchUser.CompanyName.Returns(testCompanyName);
            testExchUser.JobTitle.Returns(testJobTitle);

            // テストするメソッドにアクセスし、実際の結果を取得
            RecipientInformationDto actual = (RecipientInformationDto)mi.Invoke(obj, new object[] { testSharing });

            // 期待結果
            RecipientInformationDto expected = new RecipientInformationDto(testName, testDepartment, testCompanyName, expectedJobTitle, OlMailRecipientType.olOriginator);

            // actualとexpectedを比較
            CompareRecInfoDto(actual, expected);
        }
      
        /// <summary>
        /// <para>アイテムが、SharingItemの場合</para>
        /// <para>（送信者のAddressEntryが取得できる）</para>
        /// <para>（送信者のAddressEntryから、ExchangeUserが取得できない場合）</para>
        /// </summary>
        /// <remarks>
        /// 【期待結果】
        /// <para>senderInformationDtoのrecipientTypeとemailAddressが取得できる</para>
        /// </remarks>
        [Test]
        public void GetSenderInfoSharingTest2()
        {
            string testName = "Kosaka Kenta (小坂 健太)";
            string testEmailAddress = "kenta.kosaka@jp.ricoh.com";
            string expectedNameAndAddress = string.Format("{0}<{1}>", testName, testEmailAddress);

            testAdd = Substitute.For<AddressEntry>();

            // モックでつかうデータを用意
            testSharing.SenderName.Returns(testName);
            testSharing.SenderEmailAddress.Returns(testEmailAddress);

            // モックのReturn値を設定
            testRec.AddressEntry.Returns(testAdd);

            // GetExchangeUserメソッドで、例外が発生
            testAdd.GetExchangeUser().Returns(x => { throw new System.Exception(); });

            // テストするメソッドにアクセスし、実際の結果を取得
            RecipientInformationDto actual = (RecipientInformationDto)mi.Invoke(obj, new object[] { testSharing });

            // 期待結果
            RecipientInformationDto expected = new RecipientInformationDto(expectedNameAndAddress, OlMailRecipientType.olOriginator);

            // actualとexpectedを比較
            CompareRecInfoDto(actual, expected);
        }

        /// <summary>
        /// <para>アイテムが、SharingItemの場合</para>
        /// <para>（送信者のAddressEntryが取得できない）</para>
        /// </summary>
        /// <remarks>
        /// 【期待結果】
        /// <para>senderInformationDtoがnull</para>
        /// </remarks>
        [Test]

        public void GetSenderInfoSharingTest3()
        {
            string testEmailAddress = "kenta.kosaka@jp.ricoh.com";

            testAdd = Substitute.For<AddressEntry>();

            // モックでつかうデータを用意
            testSharing.SenderName.Returns((string)null);
            testSharing.SenderEmailAddress.Returns(testEmailAddress);

            // モックのReturn値を設定
            testRec.AddressEntry.Returns(testAdd);

            // GetExchangeUserメソッドで、例外が発生
            testAdd.GetExchangeUser().Returns(x => { throw new System.Exception(); });

            // テストするメソッドにアクセスし、実際の結果を取得
            RecipientInformationDto actual = (RecipientInformationDto)mi.Invoke(obj, new object[] { testSharing });

            // メソッドの戻り値がnullであることを確認
            Assert.IsNull(actual);
        }

        /// <summary>
        /// <para>アイテムが、DocumentItemの場合</para>
        /// </summary>
        /// <remarks>
        /// 【期待結果】
        /// <para>senderInformationDtoがnull</para>
        /// </remarks>
        [Test]

        public void GetSenderInfoDocumentTest1()
        {

            // テストするメソッドにアクセスし、実際の結果を取得
            RecipientInformationDto actual = (RecipientInformationDto)mi.Invoke(obj, new object[] { testDocument });

            // メソッドの戻り値がnullであることを確認
            Assert.IsNull(actual);
        }

        /// <summary>
        /// テスト対象メソッドの戻り値と、期待結果を比較するメソッド
        /// </summary>
        /// <param name="actual">メソッドからもどってきたRecipientInformationDto</param>
        /// <param name="expected">期待する結果を入れたRecipientInformationDto</param>
        private void CompareRecInfoDto(RecipientInformationDto actual, RecipientInformationDto expected)
        {
            Assert.That(actual.fullName, Is.EqualTo(expected.fullName));
            Assert.That(actual.division, Is.EqualTo(expected.division));
            Assert.That(actual.companyName, Is.EqualTo(expected.companyName));
            Assert.That(actual.recipientType, Is.EqualTo(expected.recipientType));
            Assert.That(actual.jobTitle, Is.EqualTo(expected.jobTitle));
            Assert.That(actual.emailAddress, Is.EqualTo(expected.emailAddress));
        }

    }

}
