using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Runtime.InteropServices;
using System.ComponentModel.Design;
using System.IO;
using System.Linq;
using System.Net;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using EnvDTE;
using EnvDTE80;
using Microsoft.SharePoint.Client;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.VisualStudio.Shell;
using OfficeDevPnP.Core.Utilities;
using File = System.IO.File;

namespace Britehouse.SPFileCopy
{
    /// <summary>
    /// This is the class that implements the package exposed by this assembly.
    ///
    /// The minimum requirement for a class to be considered a valid package for Visual Studio
    /// is to implement the IVsPackage interface and register itself with the shell.
    /// This package uses the helper classes defined inside the Managed Package Framework (MPF)
    /// to do it: it derives from the Package class that provides the implementation of the 
    /// IVsPackage interface and uses the registration attributes defined in the framework to 
    /// register itself and its components with the shell.
    /// </summary>
    // This attribute tells the PkgDef creation utility (CreatePkgDef.exe) that this class is
    // a package.
    [PackageRegistration(UseManagedResourcesOnly = true)]
    // This attribute is used to register the information needed to show this package
    // in the Help/About dialog of Visual Studio.
    [InstalledProductRegistration("#110", "#112", "1.0", IconResourceID = 400)]
    // This attribute is needed to let the shell know that this package exposes some menus.
    [ProvideMenuResource("Menus.ctmenu", 1)]
    [Guid(GuidList.guidSPFileCopyPkgString)]
    [ProvideAutoLoad(UIContextGuids80.SolutionExists)]
    // ReSharper disable once InconsistentNaming
    public sealed class SPFileCopyPackage : Package
    {
        /// <summary>
        /// Default constructor of the package.
        /// Inside this method you can place any initialization code that does not require 
        /// any Visual Studio service because at this point the package object is created but 
        /// not sited yet inside Visual Studio environment. The place to do all the other 
        /// initialization is the Initialize method.
        /// </summary>
        public SPFileCopyPackage()
        {
            Debug.WriteLine(string.Format(CultureInfo.CurrentCulture, "Entering constructor for: {0}", ToString()));
            LogToOutput("Entering constructor");
        }

        /////////////////////////////////////////////////////////////////////////////
        // Overridden Package Implementation
        #region Package Members

        private DocumentEvents _docEvents;
        private ProjectItemsEvents _projItemsEvents;
        private Web web;
        private ClientContext ctx;
        private Hashtable folders;

        /// <summary>
        /// Initialization of the package; this method is called right after the package is sited, so this is the place
        /// where you can put all the initialization code that rely on services provided by VisualStudio.
        /// </summary>
        protected override void Initialize()
        {
            Debug.WriteLine(string.Format(CultureInfo.CurrentCulture, "Entering Initialize() of: {0}", ToString()));
            LogToOutput("Entering Initialize");
            try
            {
                base.Initialize();

                // Add our command handlers for menu (commands must exist in the .vsct file)
                var mcs = GetService(typeof(IMenuCommandService)) as OleMenuCommandService;
                if (null != mcs)
                {
                    // Create the command for the menu item.
                    var menuCommandId = new CommandID(GuidList.guidSPFileCopyCmdSet, (int)PkgCmdIDList.cmdidUpdateFileReferences);
                    //var menuItem = new MenuCommand(MenuItemCallback, menuCommandId );

                    var menuItem = new OleMenuCommand(MenuItemCallback, menuCommandId);
                    menuItem.BeforeQueryStatus += menuCommand_BeforeQueryStatus;
                    mcs.AddCommand(menuItem);
                }

                var dte = (DTE2)GetService(typeof(DTE));
                var dteEvents = (Events2)dte.Events;

                _projItemsEvents = dteEvents.ProjectItemsEvents;

                _projItemsEvents.ItemRenamed += ProjectItemsEvents_ItemRenamed;
                _projItemsEvents.ItemAdded += ProjectItemsEvents_ItemAdded;
                _projItemsEvents.ItemRemoved += ProjectItemsEvents_ItemRemoved;

                //remember to declare it
                _docEvents = dteEvents.DocumentEvents; // defend against Garbage Collector

                _docEvents.DocumentSaved += DocumentEvents_DocumentSaved;
            }
            catch (Exception ex)
            {
                var error = string.Format("[Error] Exception in Initialize: {0}", ex.Message);
                LogToOutput(error);
            }
        }
        #endregion

        protected override void Dispose(bool disposing)
        {
            if (ctx != null)
            {
                ctx.Dispose();
                ctx = null;
            }
            base.Dispose(disposing);
        }

        private void menuCommand_BeforeQueryStatus(object sender, EventArgs e)
        {
            if (e == null) throw new ArgumentNullException("e");
            // get the menu that fired the event
            var menuCommand = sender as OleMenuCommand;
            if (menuCommand != null)
            {
                // start by assuming that the menu will not be shown
                menuCommand.Visible = false;
                menuCommand.Enabled = false;

                IVsHierarchy hierarchy;
                uint itemid;

                if (!IsSingleProjectItemSelection(out hierarchy, out itemid)) return;
                // Get the file path
                //string itemFullPath;
                //((IVsProject)hierarchy).GetMkDocument(itemid, out itemFullPath);
                //var transformFileInfo = new FileInfo(itemFullPath);

                menuCommand.Visible = true;
                menuCommand.Enabled = true;
            }
        }

        public static bool IsSingleProjectItemSelection(out IVsHierarchy hierarchy, out uint itemid)
        {
            hierarchy = null;
            itemid = VSConstants.VSITEMID_NIL;

            var monitorSelection = GetGlobalService(typeof(SVsShellMonitorSelection)) as IVsMonitorSelection;
            var solution = GetGlobalService(typeof(SVsSolution)) as IVsSolution;
            if (monitorSelection == null || solution == null)
            {
                return false;
            }

            var hierarchyPtr = IntPtr.Zero;
            var selectionContainerPtr = IntPtr.Zero;

            try
            {
                IVsMultiItemSelect multiItemSelect;
                var hr = monitorSelection.GetCurrentSelection(out hierarchyPtr, out itemid, out multiItemSelect, out selectionContainerPtr);

                if (ErrorHandler.Failed(hr) || hierarchyPtr == IntPtr.Zero || itemid == VSConstants.VSITEMID_NIL)
                {
                    // there is no selection
                    return false;
                }

                // multiple items are selected
                if (multiItemSelect != null) return false;

                // there is a hierarchy root node selected, thus it is not a single item inside a project

                if (itemid == VSConstants.VSITEMID_ROOT) return false;

                hierarchy = Marshal.GetObjectForIUnknown(hierarchyPtr) as IVsHierarchy;
                if (hierarchy == null) return false;

                Guid guidProjectId;

                return !ErrorHandler.Failed(solution.GetGuidOfProject(hierarchy, out guidProjectId));

                // if we got this far then there is a single project item selected
            }
            finally
            {
                if (selectionContainerPtr != IntPtr.Zero)
                {
                    Marshal.Release(selectionContainerPtr);
                }

                if (hierarchyPtr != IntPtr.Zero)
                {
                    Marshal.Release(hierarchyPtr);
                }
            }
        }

        /// <summary>
        /// This function is the callback used to execute a command when the a menu item is clicked.
        /// See the Initialize method to see how the menu item is associated to this function using
        /// the OleMenuCommandService service and the MenuCommand class.
        /// </summary>
        private void MenuItemCallback(object sender, EventArgs e)
        {
            // Get the file path

            var hs = GetFileMappings(null);
            if (hs == null)
            {
                return;
            }

            // Show a Message Box to prove we were here
            var uiShell = (IVsUIShell)GetService(typeof(SVsUIShell));
            var clsid = Guid.Empty;
            int result;
            ErrorHandler.ThrowOnFailure(uiShell.ShowMessageBox(
                       0,
                       ref clsid,
                       "SPProvisioningFileGenerator",
                       string.Format(CultureInfo.CurrentCulture, "Inside {0}.MenuItemCallback()", ToString()),
                       string.Empty,
                       0,
                       OLEMSGBUTTON.OLEMSGBUTTON_OK,
                       OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST,
                       OLEMSGICON.OLEMSGICON_INFO,
                       0,        // false
                       out result));
        }

        private List<Mapping> GetFileMappings(Project project)
        {
            var updateFiles = new List<Mapping>();

            if (project.ProjectItems.Item("SPFileCopy.config") == null)
            {
                return null;
            }

            var fileName = project.ProjectItems.Item("SPFileCopy.config").FileNames[1];

            try
            {
                var xDoc = XDocument.Load(fileName);
                if (xDoc.Root == null)
                {
                    return null;
                }

                var q = (from b in xDoc.Root.Descendants("mapping")
                         select new Mapping()
                         {
                             Site = (string)b.Element("site"),
                             Source = (string)b.Element("source"),
                             Target = (string)b.Element("target")
                         }).ToList();

                return q;
            }
            catch (Exception ex)
            {
                var error = string.Format("[Error] Exception with mapping: {0}", ex.Message);
                LogToOutput(error);
                return null;
            }
        }

        private void DocumentEvents_DocumentSaved(Document document)
        {
            System.Threading.Tasks.Task.Run(() =>
            {
                CopyFiles(document);
            });
        }

        private void ProjectItemsEvents_ItemRemoved(ProjectItem projectItem)
        {
            System.Threading.Tasks.Task.Run(() =>
            {
                //try
                //{
                //    if (projectItem.Kind != EnvDTE.Constants.vsProjectItemKindPhysicalFile)
                //    {
                //        return;
                //    }

                //    var mapping = GetMapping(projectItem.ContainingProject, projectItem.FileNames[1]);
                //    if (string.IsNullOrEmpty(mapping))
                //    {
                //        return;
                //    }

                //    File.Delete(mapping);

                //    var s = "[Event] Item removed from " +
                //            Path.GetFileName(projectItem.FileNames[1]) + " in project " + projectItem.ContainingProject.Name;
                //    LogToOutput(s);
                //}
                //catch (Exception ex)
                //{
                //    LogToOutput(string.Format("[Error] Cannot delete document: {0}: {1}", projectItem.FileNames[1], ex.Message));
                //}
            });
        }

        private void ProjectItemsEvents_ItemAdded(ProjectItem projectItem)
        {
            System.Threading.Tasks.Task.Run(() =>
            {
                //try
                //{
                //    if (projectItem.Kind != EnvDTE.Constants.vsProjectItemKindPhysicalFile)
                //    {
                //        return;
                //    }

                //    var mapping = GetMapping(projectItem.ContainingProject, projectItem.FileNames[1]);
                //    if (string.IsNullOrEmpty(mapping))
                //    {
                //        return;
                //    }

                //    CopyFile(projectItem.FileNames[1], mapping);

                //    var s = "[Event] Item added to " +
                //            Path.GetFileName(projectItem.FileNames[1]) + " in project " + projectItem.ContainingProject.Name;
                //    LogToOutput(s);
                //}
                //catch (Exception ex)
                //{
                //    LogToOutput(string.Format("[Error] Cannot add document: {0}: {1}", projectItem.FileNames[1], ex.Message));
                //}
            });
        }

        public void ProjectItemsEvents_ItemRenamed(ProjectItem projectItem, string oldName)
        {
            System.Threading.Tasks.Task.Run(() =>
            {
                //try
                //{
                //    if (projectItem.Kind != EnvDTE.Constants.vsProjectItemKindPhysicalFile)
                //    {
                //        return;
                //    }

                //    var mappingFrom = GetMapping(projectItem.ContainingProject, projectItem.FileNames[1]);
                //    if (string.IsNullOrEmpty(mappingFrom))
                //    {
                //        return;
                //    }

                //    var mappingTo = GetMapping(projectItem.ContainingProject, projectItem.FileNames[1]);
                //    if (string.IsNullOrEmpty(mappingTo))
                //    {
                //        return;
                //    }

                //    MoveFile(mappingFrom, mappingTo);

                //    var s = "[Event] Renamed " + oldName + " to " +
                //            Path.GetFileName(projectItem.FileNames[1]) + " in project " + projectItem.ContainingProject.Name;
                //    LogToOutput(s);
                //}
                //catch (Exception ex)
                //{
                //    LogToOutput(string.Format("[Error] Cannot rename document: {0}: {1}", oldName, ex.Message));
                //}
            });
        }

        private List<Mapping> GetMapping(Project project, string filePath)
        {
            string fileName;
            try
            {
                fileName = project.ProjectItems.Item("SPFileCopy.config").FileNames[1];
            }
            catch
            {
                return null;
            }

            try
            {
                var xDoc = XDocument.Load(fileName);
                if (xDoc.Root == null)
                {
                    return null;
                }

                var q = (from b in xDoc.Root.Descendants("mapping")
                         let approveElement = b.Element("approve")
                         let publishElement = b.Element("publish")
                         let checkoutElement = b.Element("checkout")
                         let siteElement = b.Element("site")
                         where siteElement != null
                         let sourceElement = b.Element("source")
                         where sourceElement != null
                         let targetElement = b.Element("target")
                         where targetElement != null
                         where filePath.ToLower().Contains(sourceElement.Value.ToLower())
                         select new Mapping()
                         {
                             Site = siteElement.Value,
                             Source = sourceElement.Value,
                             Target = targetElement.Value,
                             FilePath = filePath,
                             ProjectFullPath = project.FullName,
                             Checkout = checkoutElement != null && bool.Parse(checkoutElement.Value),
                             Publish = publishElement != null && bool.Parse(publishElement.Value),
                             Approve = approveElement != null && bool.Parse(approveElement.Value)
                         }).ToList();

                return q;
                //var results = q.OrderByDescending(o => o.Source.Length).FirstOrDefault(w => filePath.ToLower().Contains(w.Source.ToLower()));

                //if (results == null) return null;

                //var source = string.Format("{0}\\{1}", Path.GetDirectoryName(project.FullName), results.Source);
                //source = source.TrimEnd('\\');
                //var target = results.Target.TrimEnd('/');

                //filePath = ReplaceCaseInsensitive(filePath, source, target);

                //return filePath;
            }
            catch (Exception ex)
            {
                var error = string.Format("[Error] Exception with mapping: {0}", ex.Message);
                LogToOutput(error);
                return null;
            }
        }

        private void CopyFiles(Document document)
        {
            var mappings = GetMapping(document.ProjectItem.ContainingProject, document.Path);
            foreach (var mapping in mappings)
            {
                try
                {
                    CopyFile(document.FullName, mapping);
                    var s = "[Event] Document saved: " +
                            Path.GetFileName(document.FullName) + " in project " +
                            document.ProjectItem.ContainingProject.Name;
                    LogToOutput(s);
                }
                catch (Exception ex)
                {
                    LogToOutput(string.Format("[Error] Cannot save document: {0}: {1}", document.FullName,
                        ex.Message));
                }
            }
        }
        private void CopyFile(string source, Mapping mapping)
        {
            var s = "[Event] Copying: " +
                            Path.GetFileName(source);
            LogToOutput(s);

            if (ctx == null || ctx.Url != mapping.Site)
            {
                try
                {
                    ctx = new ClientContext(mapping.Site) { Credentials = GetCredentials(mapping.Site) };
                    web = ctx.Web;
                    web.EnsureProperty(w => w.ServerRelativeUrl);
                }
                catch (Exception ex)
                {
                    if (ctx != null)
                    {
                        ctx.Dispose();
                    }
                    ctx = null;
                }

                folders = new Hashtable();
            }

            Folder folder;
            if (folders.ContainsKey(mapping.FullTargetFolder))
            {
                folder = (Folder)folders[mapping.FullTargetFolder];
            }
            else
            {
                folder = web.EnsureFolder(web.RootFolder, mapping.FullTargetFolder);
                folders.Add(mapping.FullTargetFolder, folder);
            }

            var fileUrl = UrlUtility.Combine(folder.ServerRelativeUrl, Path.GetFileName(source));

            // Check if the file exists
            if (mapping.Checkout)
            {
                try
                {
                    var existingFile = web.GetFileByServerRelativeUrl(fileUrl);
                    existingFile.EnsureProperty(f => f.Exists);
                    if (existingFile.Exists)
                    {
                        web.CheckOutFile(fileUrl);
                    }
                }
                catch
                { // Swallow exception, file does not exist 
                }
            }

            folder.UploadFile(new FileInfo(source).Name, source, true);

            if (mapping.Checkout)
                web.CheckInFile(fileUrl, CheckinType.MajorCheckIn, "");

            if (mapping.Publish)
                web.PublishFile(fileUrl, "Copied from SP File Copy");

            if (mapping.Approve)
                web.ApproveFile(fileUrl, "Copied from SP File Copy");
        }

        public void MoveFile(string source, string target)
        {
            var targetFolder = Path.GetDirectoryName(target);
            if (string.IsNullOrEmpty(targetFolder))
            {
                return;
            }

            if (!Directory.Exists(targetFolder))
            {
                Directory.CreateDirectory(targetFolder);
            }

            File.Copy(source, target, true);
        }

        private void LogToOutput(string message)
        {
            var dte = (DTE2)GetService(typeof(DTE));

            var outputWindow = dte.ToolWindows.OutputWindow;
            OutputWindowPane outputWindowPane;
            try
            {
                outputWindowPane = outputWindow.OutputWindowPanes.Item("SP File Copy");
            }
            catch (Exception)
            {
                outputWindowPane = outputWindow.OutputWindowPanes.Add("SP File Copy");
            }

            outputWindowPane.Activate();
            outputWindowPane.OutputString(message + "\n");
        }

        private NetworkCredential GetCredentials(string url)
        {
            NetworkCredential creds = null;

            var connectionURI = new Uri(url);

            // Try to get the credentials by full url

            creds = CredentialManager.GetCredential(url);
            if (creds != null) return creds;

            // Try to get the credentials by splitting up the path
            var pathString = string.Format("{0}://{1}", connectionURI.Scheme, connectionURI.IsDefaultPort ? connectionURI.Host : string.Format("{0}:{1}", connectionURI.Host, connectionURI.Port));
            var path = connectionURI.AbsolutePath;
            while (path.IndexOf('/') != -1)
            {
                path = path.Substring(0, path.LastIndexOf('/'));
                if (string.IsNullOrEmpty(path)) continue;

                var pathUrl = string.Format("{0}{1}", pathString, path);
                creds = CredentialManager.GetCredential(pathUrl);
                if (creds != null)
                {
                    break;
                }
            }

            if (creds != null) return creds;
            // Try to find the credentials by schema and hostname
            creds = CredentialManager.GetCredential(connectionURI.Scheme + "://" + connectionURI.Host) ??
                    CredentialManager.GetCredential(connectionURI.Host);

            return creds;
        }
    }
}
