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
        private TestReportItem testReport;

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
            testReport = Substitute.For<TestReportItem>();

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

            // 期待結果を入れるリスト
            List<Recipient> expectedRecList = new List<Recipient>();
            expectedRecList.Add(expectedRec1);
            expectedRecList.Add(expectedRec2);

            // モックのReturn値と、期待結果のリストの値を設定
            testMail.Recipients.Count.Returns(testRecNames.Length);
            SubstituteRecProps(testRecNames, testRecSendable, testRecType, testMail);
            SetExpectedValues(testRecNames, testRecSendable, testRecType, expectedRecList);

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

            // actualとexpectedのリストを比較
            Assert.AreEqual(actualRecList.Count, expectedRecList.Count);
            CompareLists(actualRecList, expectedRecList);

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

            // 期待結果を入れるリスト
            List<Recipient> expectedRecList = new List<Recipient>();
            expectedRecList.Add(expectedRec1);
            expectedRecList.Add(expectedRec2);

            // モックのReturn値と、期待結果のリストの値を設定
            testMeeting.Recipients.Count.Returns(testRecNames.Length);
            SubstituteRecProps(testRecNames, testRecSendable, testRecType, testMeeting);
            SetExpectedValues(testRecNames, testRecSendable, testRecType, expectedRecList);
            testMeeting.MessageClass.Returns("IPM.Schedule.Meeting.Request");

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

            // actualとexpectedのリストを比較
            Assert.AreEqual(actualRecList.Count, expectedRecList.Count);
            CompareLists(actualRecList, expectedRecList);

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

            // 期待結果を入れるリスト
            List<Recipient> expectedRecList = new List<Recipient>();
            expectedRecList.Add(expectedRec1);
            expectedRecList.Add(expectedRec2);

            // モックのReturn値と、期待結果のリストの値を設定
            testMeeting.Recipients.Count.Returns(testRecNames.Length);
            testMeeting.MessageClass.Returns("IPM.Schedule.Meeting.Resp.Pos");
            SubstituteRecProps(testRecNames, testRecSendable, testRecType, testMeeting);
            SetExpectedValues(testRecNames, testRecSendable, testRecType, expectedRecList);

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

            // actualとexpectedのリストを比較
            Assert.AreEqual(actualRecList.Count, expectedRecList.Count);
            CompareLists(actualRecList, expectedRecList);

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

            // 期待結果を入れるリスト
            List<Recipient> expectedRecList = new List<Recipient>();
            expectedRecList.Add(expectedRec1);
            expectedRecList.Add(expectedRec2);

            // モックのReturn値と、期待結果のリストの値を設定
            testMeeting.Recipients.Count.Returns(testRecNames.Length);
            testMeeting.MessageClass.Returns("IPM.Schedule.Meeting.Canceled");
            SubstituteRecProps(testRecNames, testRecSendable, testRecType, testMeeting);
            SetExpectedValues(testRecNames, testRecSendable, testRecType, expectedRecList);

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

            // actualとexpectedのリストを比較
            Assert.AreEqual(actualRecList.Count, expectedRecList.Count);
            CompareLists(actualRecList, expectedRecList);

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

            // 期待結果を入れるリスト
            List<Recipient> expectedRecList = new List<Recipient>();
            expectedRecList.Add(expectedRec1);
            expectedRecList.Add(expectedRec2);

            // モックのReturn値と、期待結果のリストの値を設定
            testMeeting.Recipients.Count.Returns(testRecNames.Length);
            testMeeting.MessageClass.Returns("IPM.Schedule.Meeting.Resp.Neg");
            SubstituteRecProps(testRecNames, testRecSendable, testRecType, testMeeting);
            SetExpectedValues(testRecNames, testRecSendable, testRecType, expectedRecList);

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
            Application app = new Application();
            ExchangeUser currentUser = app.Session.CurrentUser.AddressEntry.GetExchangeUser();

            // モックでつかうデータを用意
            //（自分と、BCC(リソース)もテスト用Recに入れる）
            string[] testRecNames = { currentUser.Address, "testemailaddress1@example.com", "testemailaddress2@example.com", "testemailaddress3@example.com" };
            bool[] testRecSendable = { true, true, true, false };
            int[] testRecType = { (int)OlMailRecipientType.olOriginator, (int)OlMailRecipientType.olTo, (int)OlMailRecipientType.olBCC, (int)OlMailRecipientType.olBCC };

            // モックのReturn値と、期待結果のリストの値を設定
            testAppointment.Recipients.Count.Returns(testRecNames.Length);
            SubstituteRecProps(testRecNames, testRecSendable, testRecType, testAppointment);

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

            // リストのサイズから、自分（送信者）と、Sendableがfalseのリソースが返り値のリストに入っていないことを確認
            Assert.AreEqual(actualRecList.Count, expectedRecList.Count, testRecNames.Length - 2);
            CompareLists(actualRecList, expectedRecList);

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
            Application app = new Application();
            ExchangeUser currentUser = app.Session.CurrentUser.AddressEntry.GetExchangeUser();

            // モックでつかうデータを用意
            //（自分と、BCC(リソース)もテスト用Recに入れる）
            string[] testRecNames = { currentUser.Address, "testemailaddress1@example.com", "testemailaddress2@example.com" };
            bool[] testRecSendable = { true, true, true };
            int[] testRecType = { (int)OlMailRecipientType.olOriginator, (int)OlMailRecipientType.olTo, (int)OlMailRecipientType.olCC };

            // モックのReturn値と、期待結果のリストの値を設定
            testAppointment.Recipients.Count.Returns(testRecNames.Length);
            SubstituteRecProps(testRecNames, testRecSendable, testRecType, testAppointment);

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
            CompareLists(actualRecList, expectedRecList);

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

            // 期待結果を入れるリスト
            List<Recipient> expectedRecList = new List<Recipient>();
            expectedRecList.Add(expectedRec1);
            expectedRecList.Add(expectedRec2);

            // モックのReturn値と、期待結果のリストの値を設定
            testSharing.Recipients.Count.Returns(testRecNames.Length);
            SubstituteRecProps(testRecNames, testRecSendable, testRecType, testSharing);
            SetExpectedValues(testRecNames, testRecSendable, testRecType, expectedRecList);

            // テストするメソッドにアクセスし、実際の結果を取得
            var objArray = new object[] { testSharing, Utility.OutlookItemType.Mail, false };
            object actualObj = mi.Invoke(obj, objArray);
            // テスト対象メソッドの返り値をList<Recipient>型にする
            List<Recipient> actualRecList = new List<Recipient>();
            IEnumerable<Recipient> actualEnumList = (IEnumerable<Recipient>)actualObj;

            foreach (var actual in actualEnumList)
            {
                actualRecList.Add(actual);
            }

            // actualとexpectedのリストを比較
            Assert.AreEqual(actualRecList.Count, expectedRecList.Count);
            CompareLists(actualRecList, expectedRecList);

            // ref引数のtypeが正しいことを確認
            Assert.That(objArray[1], Is.EqualTo(Utility.OutlookItemType.Sharing));

        }

        /// <summary>
        ///  ReportItemの場合
        ///  Recipientsを取得でき、TypeがReportになる
        /// </summary>
        [Test]
        public void GetRecipientsTest9()
        {
            // モックでつかうデータを用意
            string[] testRecNames = { "testemailaddress1@example.com", "testemailaddress2@example.com" };
            bool[] testRecSendable = { true, true };
            int[] testRecType = { (int)OlMailRecipientType.olTo, (int)OlMailRecipientType.olCC };

            // 期待結果を入れるリスト
            List<Recipient> expectedRecList = new List<Recipient>();
            expectedRecList.Add(expectedRec1);
            expectedRecList.Add(expectedRec2);

            testReport.CopyHon().Returns(testReport);

            //testMail.Recipients.Count.Returns(testRecNames.Length);
            //SubstituteRecProps(testRecNames, testRecSendable, testRecType, testMail);


            // モックのReturn値と、期待結果のリストの値を設定
            MyTestNs myTestNs = Substitute.For<MyTestNs>();
            testApp.Session.Returns(myTestNs);
            myTestNs.GetItemFromIDHon(Arg.Any<string>()).Returns(testMail);

            
            // テストするメソッドにアクセスし、実際の結果を取得
            var objArray = new object[] { testReport, Utility.OutlookItemType.Mail, false };
            object actualObj = mi.Invoke(obj, objArray);

            // テスト対象メソッドの返り値をList<Recipient>型にする
            List<Recipient> actualRecList = new List<Recipient>();
            IEnumerable<Recipient> actualEnumList = (IEnumerable<Recipient>)actualObj;
            foreach (var actual in actualEnumList)
            {
                actualRecList.Add(actual);
            }
            
            SetExpectedValues(testRecNames, testRecSendable, testRecType, expectedRecList);

            // actualとexpectedのリストを比較
            Assert.AreEqual(actualRecList.Count, expectedRecList.Count);
            CompareLists(actualRecList, expectedRecList);

            // ref引数のtypeが正しいことを確認
            Assert.That(objArray[1], Is.EqualTo(Utility.OutlookItemType.Report));

        }

        /// <summary>
        /// テスト対象メソッドで使われる値のReturnsを設定するメソッド
        /// </summary>
        /// <param name="testRecNames">Recipientのアドレス</param>
        /// <param name="testRecSendable">RecipientのSendableプロパティ</param>
        /// <param name="testRecType">RecipientのType</param>
        /// <param name="item">選択されたitem</param>
        private void SubstituteRecProps(string[] testRecNames, bool[] testRecSendable, int[] testRecType, object item)
        {

            int i = 0;

            if (item is MailItem)
            {
                MailItem testItem = (MailItem)item;

                foreach (string testRec in testRecNames)
                {
                    // テスト用Recipientのプロパティに値を設定
                    testItem.Recipients[i + 1].Address.Returns(testRecNames[i]);
                    testItem.Recipients[i + 1].Sendable.Returns(testRecSendable[i]);
                    testItem.Recipients[i + 1].Type.Returns(testRecType[i]);

                    i++;
                }
            }
            else if (item is MeetingItem)
            {
                MeetingItem testItem = (MeetingItem)item;

                foreach (string testRec in testRecNames)
                {
                    // テスト用Recipientのプロパティに値を設定
                    testItem.Recipients[i + 1].Address.Returns(testRecNames[i]);
                    testItem.Recipients[i + 1].Sendable.Returns(testRecSendable[i]);
                    testItem.Recipients[i + 1].Type.Returns(testRecType[i]);

                    i++;
                }
            }
            else if (item is AppointmentItem)
            {
                AppointmentItem testItem = (AppointmentItem)item;

                foreach (string testRec in testRecNames)
                {
                    // テスト用Recipientのプロパティに値を設定
                    testItem.Recipients[i + 1].Address.Returns(testRecNames[i]);
                    testItem.Recipients[i + 1].Sendable.Returns(testRecSendable[i]);
                    testItem.Recipients[i + 1].Type.Returns(testRecType[i]);

                    i++;
                }

            }
            else if (item is SharingItem)
            {
                SharingItem testItem = (SharingItem)item;

                foreach (string testRec in testRecNames)
                {
                    // テスト用Recipientのプロパティに値を設定
                    testItem.Recipients[i + 1].Address.Returns(testRecNames[i]);
                    testItem.Recipients[i + 1].Sendable.Returns(testRecSendable[i]);
                    testItem.Recipients[i + 1].Type.Returns(testRecType[i]);

                    i++;
                }
            }
            //else if (item is ReportItem)
            //{
            //    ReportItem testItem = (ReportItem)item;

            //    foreach (string testRec in testRecNames)
            //    {
            //        // テスト用Recipientのプロパティに値を設定
            //        testItem.Recipients[i + 1].Address.Returns(testRecNames[i]);
            //        testItem.Recipients[i + 1].Sendable.Returns(testRecSendable[i]);
            //        testItem.Recipients[i + 1].Type.Returns(testRecType[i]);

            //        i++;
            //    }
            //}


        }

        /// <summary>
        /// 期待する結果リストの値を設定するメソッド
        /// </summary>
        /// <param name="testRecNames">Recipientのアドレス</param>
        /// <param name="testRecSendable">RecipientのSendableプロパティ</param>
        /// <param name="testRecType">RecipientのType</param>
        /// <param name="expectedRecList">期待結果のRecipient型リスト</param>
        private void SetExpectedValues(string[] testRecNames, bool[] testRecSendable, int[] testRecType, List<Recipient> expectedRecList)
        {
            int i = 0;
            foreach (string testRec in testRecNames)
            {
                expectedRecList[i].Address.Returns(testRecNames[i]);
                expectedRecList[i].Sendable = testRecSendable[i];
                expectedRecList[i].Type = testRecType[i];
                i++;
            }
        }

        /// <summary>
        /// 実際の値と、期待する値を比較するメソッド
        /// </summary>
        /// <param name="actualList">メソッドからもどってきたRecipient型リスト</param>
        /// <param name="expectedList">期待する結果を入れたRecipient型リスト</param>
        private void CompareLists(List<Recipient> actualList, List<Recipient> expectedList)
        {
            for (int i = 0; i < expectedList.Count; i++)
            {
                Assert.That(actualList[i].Address, Is.EqualTo(expectedList[i].Address));
                Assert.That(actualList[i].Sendable, Is.EqualTo(expectedList[i].Sendable));
                Assert.That(actualList[i].Type, Is.EqualTo(expectedList[i].Type));
            }
        }


    }
}
