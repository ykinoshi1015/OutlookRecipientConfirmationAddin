using Microsoft.Office.Interop.Outlook;
using Microsoft.Office.Tools;
using Microsoft.Office.Tools.Outlook;
using System;
using System.Collections.Generic;
using NSubstitute;
using Microsoft.Office.Tools.Ribbon;
using System.ComponentModel;
using System.Windows.Forms;
using Microsoft.Office.Core;

/// <summary>
/// Unitテスト用に作成した自作クラス
/// </summary>
namespace ORCAUnitTest
{
    class MyTestClasses
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

    /// <summary>
    /// dynamiを返すCreateItemで怒られるのの対策
    /// </summary>
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

    /// <summary>
    /// dynamicを返すCopyで怒られるのの対策
    /// </summary>
    public abstract class TestReportItem : ReportItem
    {
        public abstract Actions Actions { get; }
        public abstract Microsoft.Office.Interop.Outlook.Application Application { get; }
        public abstract Attachments Attachments { get; }
        public abstract bool AutoResolvedWinner { get; }
        public abstract string BillingInformation { get; set; }
        public abstract string Body { get; set; }
        public abstract string Categories { get; set; }
        public abstract OlObjectClass Class { get; }
        public abstract string Companies { get; set; }
        public abstract Conflicts Conflicts { get; }
        public abstract string ConversationID { get; }
        public abstract string ConversationIndex { get; }
        public abstract string ConversationTopic { get; }
        public abstract DateTime CreationTime { get; }
        public abstract OlDownloadState DownloadState { get; }
        public abstract string EntryID { get; }
        public abstract FormDescription FormDescription { get; }
        public abstract Inspector GetInspector { get; }
        public abstract OlImportance Importance { get; set; }
        public abstract bool IsConflict { get; }
        public abstract ItemProperties ItemProperties { get; }
        public abstract DateTime LastModificationTime { get; }
        public abstract Links Links { get; }
        public abstract dynamic MAPIOBJECT { get; }
        public abstract OlRemoteStatus MarkForDownload { get; set; }
        public abstract string MessageClass { get; set; }
        public abstract string Mileage { get; set; }
        public abstract bool NoAging { get; set; }
        public abstract int OutlookInternalVersion { get; }
        public abstract string OutlookVersion { get; }
        public abstract dynamic Parent { get; }
        public abstract PropertyAccessor PropertyAccessor { get; }
        public abstract DateTime RetentionExpirationDate { get; }
        public abstract string RetentionPolicyName { get; }
        public abstract bool Saved { get; }
        public abstract OlSensitivity Sensitivity { get; set; }
        public abstract NameSpace Session { get; }
        public abstract int Size { get; }
        public abstract string Subject { get; set; }
        public abstract bool UnRead { get; set; }
        public abstract UserProperties UserProperties { get; }

        public abstract event ItemEvents_10_AfterWriteEventHandler AfterWrite;
        public abstract event ItemEvents_10_AttachmentAddEventHandler AttachmentAdd;
        public abstract event ItemEvents_10_AttachmentReadEventHandler AttachmentRead;
        public abstract event ItemEvents_10_AttachmentRemoveEventHandler AttachmentRemove;
        public abstract event ItemEvents_10_BeforeAttachmentAddEventHandler BeforeAttachmentAdd;
        public abstract event ItemEvents_10_BeforeAttachmentPreviewEventHandler BeforeAttachmentPreview;
        public abstract event ItemEvents_10_BeforeAttachmentReadEventHandler BeforeAttachmentRead;
        public abstract event ItemEvents_10_BeforeAttachmentSaveEventHandler BeforeAttachmentSave;
        public abstract event ItemEvents_10_BeforeAttachmentWriteToTempFileEventHandler BeforeAttachmentWriteToTempFile;
        public abstract event ItemEvents_10_BeforeAutoSaveEventHandler BeforeAutoSave;
        public abstract event ItemEvents_10_BeforeCheckNamesEventHandler BeforeCheckNames;
        public abstract event ItemEvents_10_BeforeDeleteEventHandler BeforeDelete;
        public abstract event ItemEvents_10_BeforeReadEventHandler BeforeRead;
        public abstract event ItemEvents_10_CustomActionEventHandler CustomAction;
        public abstract event ItemEvents_10_CustomPropertyChangeEventHandler CustomPropertyChange;
        public abstract event ItemEvents_10_ForwardEventHandler Forward;
        public abstract event ItemEvents_10_OpenEventHandler Open;
        public abstract event ItemEvents_10_PropertyChangeEventHandler PropertyChange;
        public abstract event ItemEvents_10_ReadEventHandler Read;
        public abstract event ItemEvents_10_ReadCompleteEventHandler ReadComplete;
        public abstract event ItemEvents_10_ReplyEventHandler Reply;
        public abstract event ItemEvents_10_ReplyAllEventHandler ReplyAll;
        public abstract event ItemEvents_10_SendEventHandler Send;
        public abstract event ItemEvents_10_UnloadEventHandler Unload;
        public abstract event ItemEvents_10_WriteEventHandler Write;

        event ItemEvents_10_CloseEventHandler ItemEvents_10_Event.Close
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

        public abstract void Close(OlInspectorClose SaveMode);

        public dynamic Copy()
        {
            return CopyHon();
        }

        public virtual ReportItem CopyHon()
        {
            return null;
        }

        ///// <summary>
        ///// テスト対象のクラス(e.g. Office365Contact)が実行するクラス
        ///// </summary>
        ///// <param name="ItemType">生成するアイテムのタイプ</param>
        ///// <returns></returns>
        //public dynamic CreateItem(OlItemType ItemType)
        //{
        //    return CreateItemHon(ItemType);
        //}

        ///// <summary>
        ///// dynamicを返さないために作ったメソッド
        ///// virtualをつけることで、Substituteに制御を渡す
        ///// </summary>
        ///// <param name="ItemType">生成するアイテムのタイプ</param>
        ///// <returns></returns>
        //public virtual ContactItem CreateItemHon(OlItemType ItemType)
        //{
        //    return null;
        //}


        public abstract void Delete();
        public abstract void Display(object Modal);
        public abstract Conversation GetConversation();
        public abstract dynamic Move(MAPIFolder DestFldr);
        public abstract void PrintOut();
        public abstract void Save();
        public abstract void SaveAs(string Path, object Type);
        public abstract void ShowCategoriesDialog();
    }

    /// <summary>
    /// dynamicを返すGetItemFromIDで怒られるのの対策
    /// </summary>
    public abstract class MyTestNs : NameSpace
    {
        public abstract Accounts Accounts { get; }
        public abstract AddressLists AddressLists { get; }
        public abstract Microsoft.Office.Interop.Outlook.Application Application { get; }
        public abstract OlAutoDiscoverConnectionMode AutoDiscoverConnectionMode { get; }
        public abstract string AutoDiscoverXml { get; }
        public abstract Categories Categories { get; }
        public abstract OlObjectClass Class { get; }
        public abstract string CurrentProfileName { get; }
        public abstract Recipient CurrentUser { get; }
        public abstract Store DefaultStore { get; }
        public abstract OlExchangeConnectionMode ExchangeConnectionMode { get; }
        public abstract string ExchangeMailboxServerName { get; }
        public abstract string ExchangeMailboxServerVersion { get; }
        public abstract Folders Folders { get; }
        public abstract dynamic MAPIOBJECT { get; }
        public abstract bool Offline { get; }
        public abstract dynamic Parent { get; }
        public abstract NameSpace Session { get; }
        public abstract Stores Stores { get; }
        public abstract SyncObjects SyncObjects { get; }
        public abstract string Type { get; }

        public abstract event NameSpaceEvents_OptionsPagesAddEventHandler OptionsPagesAdd;
        public abstract event NameSpaceEvents_AutoDiscoverCompleteEventHandler AutoDiscoverComplete;

        public abstract void AddStore(object Store);
        public abstract void AddStoreEx(object Store, OlStoreType Type);
        public abstract bool CompareEntryIDs(string FirstEntryID, string SecondEntryID);
        public abstract ContactCard CreateContactCard(AddressEntry AddressEntry);
        public abstract Recipient CreateRecipient(string RecipientName);
        public abstract SharingItem CreateSharingItem(object Context, object Provider);
        public abstract void Dial(object ContactItem);
        public abstract AddressEntry GetAddressEntryFromID(string ID);
        public abstract MAPIFolder GetDefaultFolder(OlDefaultFolders FolderType);
        public abstract MAPIFolder GetFolderFromID(string EntryIDFolder, object EntryIDStore);
        public abstract AddressList GetGlobalAddressList();
        //public abstract dynamic GetItemFromID(string EntryIDItem, object EntryIDStore);


        public dynamic GetItemFromID(string EntryIDItem)
        {
            return GetItemFromIDHon(EntryIDItem);
        }

        public virtual MailItem GetItemFromIDHon(string EntryIDItem)
        {
            return null;
        }


        public dynamic GetItemFromID(string EntryIDItem, object EntryIDStore)
        {
            MailItem mockMail = Substitute.For<MailItem>();

            // モックでつかうデータを用意
            string[] testRecNames = { "testemailaddress1@example.com", "testemailaddress2@example.com" };
            bool[] testRecSendable = { true, true };
            int[] testRecType = { (int)OlMailRecipientType.olTo, (int)OlMailRecipientType.olCC };

            int i = 0;

            foreach (string testRec in testRecNames)
            {
                mockMail.Recipients[i + 1].Address.Returns(testRecNames[i]);
                mockMail.Recipients[i + 1].Sendable.Returns(testRecSendable[i]);
                mockMail.Recipients[i + 1].Type.Returns(testRecType[i]);
                i++;
            }
            mockMail.Recipients.Count.Returns(testRecNames.Length);
            return mockMail;

           // return GetItemFromIDHon(EntryIDItem, EntryIDStore);
        }

        public virtual MailItem GetItemFromIDHon(string EntryIDItem, object EntryIDStore)
        {

            MailItem mockMail = Substitute.For<MailItem>();

            //// モックでつかうデータを用意
            //string[] testRecNames = { "testemailaddress1@example.com", "testemailaddress2@example.com" };
            //bool[] testRecSendable = { true, true };
            //int[] testRecType = { (int)OlMailRecipientType.olTo, (int)OlMailRecipientType.olCC };

            //int i = 0;

            //foreach (string testRec in testRecNames)
            //{
            //    mockMail.Recipients[i + 1].Address.Returns(testRecNames[i]);
            //    mockMail.Recipients[i + 1].Sendable.Returns(testRecSendable[i]);
            //    mockMail.Recipients[i + 1].Type.Returns(testRecType[i]);
            //    i++;
            //}
            return mockMail;
        }

        public abstract Recipient GetRecipientFromID(string EntryID);
        public abstract SelectNamesDialog GetSelectNamesDialog();
        public abstract MAPIFolder GetSharedDefaultFolder(Recipient Recipient, OlDefaultFolders FolderType);
        public abstract Store GetStoreFromID(string ID);
        public abstract void Logoff();
        public abstract void Logon(object Profile, object Password, object ShowDialog, object NewSession);
        public abstract MAPIFolder OpenSharedFolder(string Path, object Name, object DownloadAttachments, object UseTTL);
        public abstract dynamic OpenSharedItem(string Path);
        public abstract MAPIFolder PickFolder();
        public abstract void RefreshRemoteHeaders();
        public abstract void RemoveStore(MAPIFolder Folder);
        public abstract void SendAndReceive(bool showProgressDialog);
    }
}
