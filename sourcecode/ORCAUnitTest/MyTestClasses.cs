using Microsoft.Office.Interop.Outlook;
using Microsoft.Office.Tools;
using Microsoft.Office.Tools.Outlook;
using System;
using System.Collections.Generic;
using NSubstitute;
using Microsoft.Office.Tools.Ribbon;
using System.ComponentModel;
using System.Windows.Forms;

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
