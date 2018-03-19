using System;
using NUnit.Framework;
using NSubstitute;
using Microsoft.Office.Interop.Outlook;
using OutlookRecipientConfirmationAddin;
using System.Reflection;
using System.Collections.Generic;
using Outlook = Microsoft.Office.Interop.Outlook;
namespace ORCAUnitTest
{
    /// <summary>
    /// UtilityクラスGetRecipientsメソッドのテストクラス
    /// </summary>
    [TestFixture]
    public class GetRecipientsUnitTest
    {
        private TestContext testContextInstance;

        /// <summary>
        ///現在のテストの実行についての情報および機能を
        ///提供するテスト コンテキストを取得または設定します。
        ///</summary>
        public TestContext TestContext
        {
            get
            {
                return testContextInstance;
            }
            set
            {
                testContextInstance = value;
            }
        }

        #region 追加のテスト属性
        //
        // テストを作成する際には、次の追加属性を使用できます:
        //
        // クラス内で最初のテストを実行する前に、ClassInitialize を使用してコードを実行してください
        // [ClassInitialize()]
        // public static void MyClassInitialize(TestContext testContext) { }
        //
        // クラス内のテストをすべて実行したら、ClassCleanup を使用してコードを実行してください
        // [ClassCleanup()]
        // public static void MyClassCleanup() { }
        //
        // 各テストを実行する前に、TestInitialize を使用してコードを実行してください
        // [TestInitialize()]
        // public void MyTestInitialize() { }
        //
        // 各テストを実行した後に、TestCleanup を使用してコードを実行してください
        // [TestCleanup()]
        // public void MyTestCleanup() { }
        //
        #endregion

        private Recipient testRec;
        private AddressEntry testAdd;
        private ExchangeUser testExchUser;
        private Module mod;
        private Type typeThisAddIn;
        private Application testApp;
        private object obj;
        private MethodInfo mi;
        private NameSpace testNs;

        private MailItem testMail;
        private MeetingItem testMeeting;
        private AppointmentItem testAppointment;
        private SharingItem testSharing;
        private Recipient expectedRec1;
        private Recipient expectedRec2; 

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

            // テスト用のXXXItemを、モックで作成
            testMail = Substitute.For<MailItem>();
            testMeeting = Substitute.For<MeetingItem>();
            testAppointment = Substitute.For<AppointmentItem>();
            testSharing = Substitute.For<SharingItem>();

            expectedRec1 = Substitute.For<Recipient>();
            expectedRec2 = Substitute.For<Recipient>();

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

            // テスト対象のクラス（Utility）のタイプを取得
            Type type = mod.GetType("OutlookRecipientConfirmationAddin.Utility");
            // インスタンスを生成し、メソッドにアクセスできるようにする
            obj = Activator.CreateInstance(type);
            // mi2 = type2.GetMethod("GetRecipients", new Type[] { typeof(object), typeof(Utility.OutlookItemType), typeof(bool)  });
            mi = type.GetMethod("GetRecipients");
        }

        /// <summary>
        ///  MailItemの場合
        ///  Recipientsを取得でき、TypeがMailのままになる
        /// </summary>
        [Test]
        public void GetRecipientsTest1()
        {


            // モックでつかうデータを用意
            string[] testRecNames = { "testemailaddress1@example.com", "testemailaddress2@example.com" };
            bool[] testRecSendable = { true, true };
            int[] testRecType = { (int)OlMailRecipientType.olTo, (int)OlMailRecipientType.olCC };

            // モックのReturn値を設定
            testMail.Recipients.Count.Returns(testRecNames.Length);
            SubstituteRecProps(testRecNames, testRecSendable, testRecType);


            // テストするメソッドにアクセスし、実際の結果を取得
            // ここではList<Recipient>にキャストできない（理由は？）
            var objArray = new object[] { testMail, Utility.OutlookItemType.Mail, false };
            object actualObj = mi.Invoke(obj, objArray);

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
            expectedRec1.Address.Returns("testemailaddress1@example.com");
            expectedRec1.Sendable.Returns(true);
            expectedRec1.Type.Returns((int)OlMailRecipientType.olTo);
            expectedRecList.Add(expectedRec1);

            // 期待結果2のデータをリストに追加
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
        ///  ItemがMeetingItemの場合（会議出席依頼の返信でない）
        ///  Recipientsを取得でき、TypeがMeetingになる
        /// </summary>
        [Test]
        public void GetRecipientsTest2()
        {


            // モックでつかうデータを用意
            string[] testRecNames = { "testemailaddress1@example.com", "testemailaddress2@example.com" };
            bool[] testRecSendable = { true, true };
            int[] testRecType = { (int)OlMailRecipientType.olTo, (int)OlMailRecipientType.olCC };

            // モックのReturn値を設定
            testMeeting.Recipients.Count.Returns(testRecNames.Length);
            testMeeting.MessageClass.Returns("IPM.Schedule.Meeting.Request");
            
            SubstituteRecProps(testRecNames, testRecSendable, testRecType);

            // テストするメソッドにアクセスし、実際の結果を取得
            var objArray = new object[] { testMeeting, Utility.OutlookItemType.Mail, false };
            object actualObj = mi.Invoke(obj, objArray);

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
            expectedRec1.Address.Returns("testemailaddress1@example.com");
            expectedRec1.Sendable.Returns(true);
            expectedRec1.Type.Returns((int)OlMailRecipientType.olTo);
            expectedRecList.Add(expectedRec1);

            // 期待結果2のデータをリストに追加
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

            // ref引数のtypeがMeetingになっていることを確認
            Assert.That(objArray[1], Is.EqualTo(Utility.OutlookItemType.Meeting));
        }

        /// <summary>
        ///  ItemがMeetingItemの場合
        ///  会議出席依頼の返信（MessageCLassに"IPM.Schedule.Meeting.Resp."が含まれる）で、IgnoreMeetingResponseがfalse
        ///  Recipientsを取得でき、TypeがMeetingResponseになる
        /// </summary>
        [Test]
        public void GetRecipientsTest3()
        {

            // モックでつかうデータを用意
            string[] testRecNames = { "testemailaddress1@example.com", "testemailaddress2@example.com" };
            bool[] testRecSendable = { true, true };
            int[] testRecType = { (int)OlMailRecipientType.olTo, (int)OlMailRecipientType.olCC };

            // モックのReturn値を設定
            testMeeting.Recipients.Count.Returns(testRecNames.Length);
            testMeeting.MessageClass.Returns("IPM.Schedule.Meeting.Resp.Pos");
            
            SubstituteRecProps(testRecNames, testRecSendable, testRecType);

            // テストするメソッドにアクセスし、実際の結果を取得
            var objArray = new object[] { testMeeting, Utility.OutlookItemType.Mail, false };
            object actualObj = mi.Invoke(obj, objArray);

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
            expectedRec1.Address.Returns("testemailaddress1@example.com");
            expectedRec1.Sendable.Returns(true);
            expectedRec1.Type.Returns((int)OlMailRecipientType.olTo);
            expectedRecList.Add(expectedRec1);

            // 期待結果2のデータをリストに追加
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

            // ref引数のtypeがMeetingになっていることを確認
            Assert.That(objArray[1], Is.EqualTo(Utility.OutlookItemType.MeetingResponse));
        }


        /// <summary>
        ///  ItemがMeetingItemの場合
        ///  会議キャンセル通知メールなど（MessageCLassに"IPM.Schedule.Meeting.Resp."が含まれない）で、IgnoreMeetingResponseがtrue
        ///  Recipientsを取得し、TypeがMeetingになる
        /// </summary>
        [Test]
        public void GetRecipientsTest4()
        {

            // モックでつかうデータを用意
            string[] testRecNames = { "testemailaddress1@example.com", "testemailaddress2@example.com" };
            bool[] testRecSendable = { true, true };
            int[] testRecType = { (int)OlMailRecipientType.olTo, (int)OlMailRecipientType.olCC };

            // モックのReturn値を設定
            testMeeting.Recipients.Count.Returns(testRecNames.Length);
            testMeeting.MessageClass.Returns("IPM.Schedule.Meeting.Canceled");


            SubstituteRecProps(testRecNames, testRecSendable, testRecType);

            // テストするメソッドにアクセスし、実際の結果を取得
            var objArray = new object[] { testMeeting, Utility.OutlookItemType.Mail, false };
            object actualObj = mi.Invoke(obj, objArray);

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
            expectedRec1.Address.Returns("testemailaddress1@example.com");
            expectedRec1.Sendable.Returns(true);
            expectedRec1.Type.Returns((int)OlMailRecipientType.olTo);
            expectedRecList.Add(expectedRec1);

            // 期待結果2のデータをリストに追加
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

            // ref引数のtypeがMeetingになっていることを確認
            Assert.That(objArray[1], Is.EqualTo(Utility.OutlookItemType.Meeting));
        }

        /// <summary>
        ///  ItemがMeetingItemの場合
        ///  会議招集メールの返信（MessageCLassに"IPM.Schedule.Meeting.Resp."が含まれる）で、IgnoreMeetingResponseがtrue
        ///  Recipientsを取得せず、TypeがMeetingResponseになる
        /// </summary>
        [Test]
        public void GetRecipientsTest5()
        {

            // モックでつかうデータを用意
            string[] testRecNames = { "testemailaddress1@example.com", "testemailaddress2@example.com" };
            bool[] testRecSendable = { true, true };
            int[] testRecType = { (int)OlMailRecipientType.olTo, (int)OlMailRecipientType.olCC };

            // モックのReturn値を設定
            testMeeting.Recipients.Count.Returns(testRecNames.Length);
            testMeeting.MessageClass.Returns("IPM.Schedule.Meeting.Resp.Neg");

            SubstituteRecProps(testRecNames, testRecSendable, testRecType);

            // テストするメソッドにアクセスし、実際の結果を取得
            var objArray = new object[] { testMeeting, Utility.OutlookItemType.Mail, true };
            object actualObj = mi.Invoke(obj, objArray);

            // メソッドの戻り値がnullであることを確認
            Assert.IsNull(actualObj);

            // ref引数のtypeがAppointmentになっていることを確認
            Assert.That(objArray[1], Is.EqualTo(Utility.OutlookItemType.MeetingResponse));
        }

        /// <summary>
        ///  ItemがAppointmentItemの場合（リソース選択あり）
        ///  自分（送信者）以外のRecipientsを取得し、TypeがAppointmentになる
        /// </summary>
        [Test]
        public void GetRecipientsTest6()
        {
         

            // 自分の情報を取得
            Outlook.Application app = new Outlook.Application();
            ExchangeUser currentUser = app.Session.CurrentUser.AddressEntry.GetExchangeUser();

            // モックでつかうデータを用意
            //（自分と、BCC(リソース)もテスト用Recに入れる）
            string[] testRecNames = { currentUser.Address, "testemailaddress1@example.com", "testemailaddress2@example.com", "testemailaddress3@example.com" };
            bool[] testRecSendable = { true, true, true, false };
            int[] testRecType = { (int)OlMailRecipientType.olOriginator, (int)OlMailRecipientType.olTo, (int)OlMailRecipientType.olBCC, (int)OlMailRecipientType.olBCC };

            // モックのReturn値を設定
            testAppointment.Recipients.Count.Returns(testRecNames.Length);


            SubstituteRecProps(testRecNames, testRecSendable, testRecType);

            // テストするメソッドにアクセスし、実際の結果を取得
            var objArray = new object[] { testAppointment, Utility.OutlookItemType.Mail, false };
            object actualObj = mi.Invoke(obj, objArray);

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
            expectedRec1.Address.Returns("testemailaddress1@example.com");
            expectedRec1.Sendable.Returns(true);
            expectedRec1.Type.Returns((int)OlMailRecipientType.olTo);
            expectedRecList.Add(expectedRec1);

            // 期待結果2のデータをリストに追加
            expectedRec2.Address.Returns("testemailaddress2@example.com");
            expectedRec2.Sendable.Returns(true);
            expectedRec2.Type.Returns((int)OlMailRecipientType.olBCC);
            expectedRecList.Add(expectedRec2);


            // リストのサイズから、自分（送信者）と、Sendableがfalseのリソースが返り値のリストに入っていないことを確認
            Assert.AreEqual(actualRecList.Count, expectedRecList.Count, testRecNames.Length - 2);

            // actualとexpectedのリストを比較
            Assert.That(actualRecList[0].Address, Is.EqualTo(expectedRecList[0].Address));
            Assert.That(actualRecList[0].Sendable, Is.EqualTo(expectedRecList[0].Sendable));
            Assert.That(actualRecList[0].Type, Is.EqualTo(expectedRecList[0].Type));

            Assert.That(actualRecList[1].Address, Is.EqualTo(expectedRecList[1].Address));
            Assert.That(actualRecList[1].Sendable, Is.EqualTo(expectedRecList[1].Sendable));
            Assert.That(actualRecList[1].Type, Is.EqualTo(expectedRecList[1].Type));

            // ref引数のtypeがAppointmentになっていることを確認
            Assert.That(objArray[1], Is.EqualTo(Utility.OutlookItemType.Appointment));
        }

        /// <summary>
        ///  ItemがAppointmentItemの場合（リソース選択なし）
        ///  自分（送信者）以外のRecipientsを取得し、TypeがAppointmentになる
        /// </summary>
        [Test]
        public void GetRecipientsTest7()
        {
            // 自分の情報を取得
            Outlook.Application app = new Outlook.Application();
            ExchangeUser currentUser = app.Session.CurrentUser.AddressEntry.GetExchangeUser();

            // モックでつかうデータを用意
            //（自分と、BCC(リソース)もテスト用Recに入れる）
            string[] testRecNames = { currentUser.Address, "testemailaddress1@example.com", "testemailaddress2@example.com" };
            bool[] testRecSendable = { true, true, true };
            int[] testRecType = { (int)OlMailRecipientType.olOriginator, (int)OlMailRecipientType.olTo, (int)OlMailRecipientType.olCC };

            // モックのReturn値を設定
            testAppointment.Recipients.Count.Returns(testRecNames.Length);


            SubstituteRecProps(testRecNames, testRecSendable, testRecType);

            // テストするメソッドにアクセスし、実際の結果を取得
            var objArray = new object[] { testAppointment, Utility.OutlookItemType.Mail, false };
            object actualObj = mi.Invoke(obj, objArray);

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
            expectedRec1.Address.Returns("testemailaddress1@example.com");
            expectedRec1.Sendable.Returns(true);
            expectedRec1.Type.Returns((int)OlMailRecipientType.olTo);
            expectedRecList.Add(expectedRec1);

            // 期待結果2のデータをリストに追加
            expectedRec2.Address.Returns("testemailaddress2@example.com");
            expectedRec2.Sendable.Returns(true);
            expectedRec2.Type.Returns((int)OlMailRecipientType.olCC);
            expectedRecList.Add(expectedRec2);

            // リストのサイズから、自分（送信者）が返り値のリストに入っていないことを確認
            Assert.AreEqual(actualRecList.Count, expectedRecList.Count, testRecNames.Length - 1);

            // actualとexpectedのリストを比較
            Assert.That(actualRecList[0].Address, Is.EqualTo(expectedRecList[0].Address));
            Assert.That(actualRecList[0].Sendable, Is.EqualTo(expectedRecList[0].Sendable));
            Assert.That(actualRecList[0].Type, Is.EqualTo(expectedRecList[0].Type));

            Assert.That(actualRecList[1].Address, Is.EqualTo(expectedRecList[1].Address));
            Assert.That(actualRecList[1].Sendable, Is.EqualTo(expectedRecList[1].Sendable));
            Assert.That(actualRecList[1].Type, Is.EqualTo(expectedRecList[1].Type));

            // ref引数のtypeがMeetingになっていることを確認
            Assert.That(objArray[1], Is.EqualTo(Utility.OutlookItemType.Appointment));
        }

        /// <summary>
        ///  SharingItemの場合
        ///  Recipientsを取得でき、TypeがSharingのままになる
        /// </summary>
        [Test]
        public void GetRecipientsTest8()
        {

            // モックでつかうデータを用意
            string[] testRecNames = { "testemailaddress1@example.com", "testemailaddress2@example.com" };
            bool[] testRecSendable = { true, true };
            int[] testRecType = { (int)OlMailRecipientType.olTo, (int)OlMailRecipientType.olCC };

            // モックのReturn値を設定
            testSharing.Recipients.Count.Returns(testRecNames.Length);
            
            SubstituteRecProps(testRecNames, testRecSendable, testRecType);

            // テストするメソッドにアクセスし、実際の結果を取得
            // ここではList<Recipient>にキャストできない（理由は？）
            var objArray = new object[] { testSharing, Utility.OutlookItemType.Mail, false };
            object actualObj = mi.Invoke(obj, objArray);

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
            expectedRec1.Address.Returns("testemailaddress1@example.com");
            expectedRec1.Sendable.Returns(true);
            expectedRec1.Type.Returns((int)OlMailRecipientType.olTo);
            expectedRecList.Add(expectedRec1);

            // 期待結果2のデータをリストに追加
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
            Assert.That(objArray[1], Is.EqualTo(Utility.OutlookItemType.Sharing));

        }

        ///// <summary>
        /////  MailItemの場合
        /////  Recipientsを取得でき、TypeがMailのままになる
        ///// </summary>
        //[Test]
        //public void GetRecipientsTest1()
        //{
        //    // テスト用のMailItemを、モックで作成
        //    MailItem testMail = Substitute.For<MailItem>();

        //    // モックでつかうデータを用意
        //    string[] testRecNames = { "testemailaddress1@example.com", "testemailaddress2@example.com" };
        //    bool[] testRecSendable = { true, true };
        //    int[] testRecType = { (int)OlMailRecipientType.olTo, (int)OlMailRecipientType.olCC };

        //    // モックのReturn値を設定
        //    testMail.Recipients.Count.Returns(testRecNames.Length);

        //    int i = 0;
        //    foreach (string testRec in testRecNames)
        //    {
        //        testMail.Recipients[i + 1].Address.Returns(testRecNames[i]);
        //        testMail.Recipients[i + 1].Sendable.Returns(testRecSendable[i]);
        //        testMail.Recipients[i + 1].Type.Returns(testRecType[i]);
        //        i++;
        //    }

        //    // テストするメソッドにアクセスし、実際の結果を取得
        //    // ここではList<Recipient>にキャストできない（理由は？）
        //    var objArray = new object[] { testMail, Utility.OutlookItemType.Mail, false };
        //    object actualObj = mi.Invoke(obj, objArray);

        //    // テスト対象メソッドの返り値をList<Recipient>型にする
        //    List<Recipient> actualRecList = new List<Recipient>();
        //    IEnumerable<Recipient> actualEnumList = (IEnumerable<Recipient>)actualObj;

        //    foreach (var actual in actualEnumList)
        //    {
        //        actualRecList.Add(actual);
        //    }

        //    // 期待結果を入れるリスト
        //    List<Recipient> expectedRecList = new List<Recipient>();

        //    // 期待結果1のデータをリストに追加
        //    Recipient expectedRec1 = Substitute.For<Recipient>();
        //    expectedRec1.Address.Returns("testemailaddress1@example.com");
        //    expectedRec1.Sendable.Returns(true);
        //    expectedRec1.Type.Returns((int)OlMailRecipientType.olTo);
        //    expectedRecList.Add(expectedRec1);

        //    // 期待結果2のデータをリストに追加
        //    Recipient expectedRec2 = Substitute.For<Recipient>();
        //    expectedRec2.Address.Returns("testemailaddress2@example.com");
        //    expectedRec2.Sendable.Returns(true);
        //    expectedRec2.Type.Returns((int)OlMailRecipientType.olCC);
        //    expectedRecList.Add(expectedRec2);

        //    // actualとexpectedのリストを比較
        //    Assert.AreEqual(actualRecList.Count, expectedRecList.Count);

        //    Assert.That(actualRecList[0].Address, Is.EqualTo(expectedRecList[0].Address));
        //    Assert.That(actualRecList[0].Sendable, Is.EqualTo(expectedRecList[0].Sendable));
        //    Assert.That(actualRecList[0].Type, Is.EqualTo(expectedRecList[0].Type));

        //    Assert.That(actualRecList[1].Address, Is.EqualTo(expectedRecList[1].Address));
        //    Assert.That(actualRecList[1].Sendable, Is.EqualTo(expectedRecList[1].Sendable));
        //    Assert.That(actualRecList[1].Type, Is.EqualTo(expectedRecList[1].Type));

        //    // ref引数のtypeが正しいことを確認
        //    Assert.That(objArray[1], Is.EqualTo(Utility.OutlookItemType.Mail));

        //}

        private void SubstituteRecProps(string[] testRecNames, bool[] testRecSendable, int[] testRecType)
        {
            int i = 0;
            foreach (string testRec in testRecNames)
            {
                testMail.Recipients[i + 1].Address.Returns(testRecNames[i]);
                testMail.Recipients[i + 1].Sendable.Returns(testRecSendable[i]);
                testMail.Recipients[i + 1].Type.Returns(testRecType[i]);
                i++;
            }
        }
    }
}
