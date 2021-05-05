using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System;
using System.Collections.Generic;
using System.ComponentModel.Design;
using System.Globalization;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using Task = System.Threading.Tasks.Task;


namespace LeoReportExtension
{
    [ProvideAutoLoad(Microsoft.VisualStudio.Shell.Interop.UIContextGuids80.SolutionExists)]
    [ProvideAutoLoad(VSConstants.UICONTEXT.SolutionExists_string)]
    [ProvideAutoLoad(UIContextGuids80.SolutionExists)]
    /// <summary>
    /// Command handler
    /// </summary>
    internal sealed class Command1
    {
        #region Declarations
        /// <summary>
        /// Command ID.
        /// </summary>
        public const int CommandId = 0x0100;

        /// <summary>
        /// Command menu group (command set GUID).
        /// </summary>
        public static readonly Guid CommandSet = new Guid("fc37e2ad-90b0-4267-99fa-0ec37b86f0c5");

        /// <summary>
        /// VS Package that provides this command, not null.
        /// </summary>
        private readonly AsyncPackage package;
        #endregion

        #region Property
        /// <summary>
        /// Gets the instance of the command.
        /// </summary>
        public static Command1 Instance
        {
            get;
            private set;
        }

        /// <summary>
        /// Gets the service provider from the owner package.
        /// </summary>
        private Microsoft.VisualStudio.Shell.IAsyncServiceProvider ServiceProvider
        {
            get
            {
                return this.package;
            }
        }
        #endregion

        #region Memberfunction
        /// <summary>
        /// Initializes a new instance of the <see cref="Command1"/> class.
        /// Adds our command handlers for menu (commands must exist in the command table file)
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        /// <param name="commandService">Command service to add command to, not null.</param>
        private Command1(AsyncPackage package, OleMenuCommandService commandService)
        {
            this.package = package ?? throw new ArgumentNullException(nameof(package));
            commandService = commandService ?? throw new ArgumentNullException(nameof(commandService));

            var menuCommandID = new CommandID(CommandSet, CommandId);
            var menuItem = new MenuCommand(this.Execute, menuCommandID);
            commandService.AddCommand(menuItem);
            DTE2 dte2 = (DTE2)Package.GetGlobalService(typeof(SDTE));
            dte2.Events.BuildEvents.OnBuildDone += BuildEvents_OnBuildDone1;
        }

        /// <summary>
        /// 編譯完成事件
        /// </summary>
        /// <param name="Scope"></param>
        /// <param name="Action"></param>
        private void BuildEvents_OnBuildDone1(vsBuildScope Scope, vsBuildAction Action)
        {
            RemoveTagAutomatically();
        }

        private void MenuItemCallback(object sender, EventArgs e)
        {
            DTE2 dte = (DTE2)Package.GetGlobalService(typeof(DTE));
            OutputWindow outputWindow = dte.ToolWindows.OutputWindow;

            OutputWindowPane outputWindowPane = outputWindow.OutputWindowPanes.Add("A New Pane");
            outputWindowPane.OutputString("Some Text");
        }


        /// <summary>
        /// Initializes the singleton instance of the command.
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        public static async Task InitializeAsync(AsyncPackage package)
        {
            // Switch to the main thread - the call to AddCommand in Command1's constructor requires
            // the UI thread.
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(package.DisposalToken);

            OleMenuCommandService commandService = await package.GetServiceAsync(typeof(IMenuCommandService)) as OleMenuCommandService;
            Instance = new Command1(package, commandService);
        }

        /// <summary>
        /// This function is the callback used to execute the command when the menu item is clicked.
        /// See the constructor to see how the menu item is associated with this function using
        /// OleMenuCommandService service and MenuCommand class.
        /// </summary>
        /// <param name="sender">Event sender.</param>
        /// <param name="e">Event args.</param>
        private void Execute(object sender, EventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            VsShellUtilities.ShowMessageBox(
                this.package,
                "",
                "Automatically Remove RDLC Report Running",
                OLEMSGICON.OLEMSGICON_INFO,
                OLEMSGBUTTON.OLEMSGBUTTON_OK,
                OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);
        }

        private void RemoveTagAutomatically()
        {
            string removeFileList = string.Empty;
            ThreadHelper.ThrowIfNotOnUIThread();
            string docName = GetActiveProjectPath();
            foreach (var item in Directory.GetFiles(docName))
            {
                if (item.Contains("rdlc"))
                {
                    StreamReader str = new StreamReader(item);
                    var data = str.ReadToEnd();
                    str.Close();
                    //replace 2016 to 2008
                    data = data.Replace("http://schemas.microsoft.com/sqlserver/reporting/2016/01/reportdefinition", "http://schemas.microsoft.com/sqlserver/reporting/2008/01/reportdefinition");

                    //remove ReportSections
                    data = data.Replace("<ReportSections>", "");
                    data = data.Replace("</ReportSections>", "");

                    //remove ReportSection
                    data = data.Replace("<ReportSection>", "");
                    data = data.Replace("</ReportSection>", "");

                    //remove ReportParametersLayout
                    var removeLayoutTagStartIndex = data.IndexOf("<ReportParametersLayout>") + 1;
                    var removeLayoutTagEndIndex = data.IndexOf("</ReportParametersLayout>") + 1;
                    data = data.Remove(removeLayoutTagStartIndex, removeLayoutTagEndIndex - removeLayoutTagStartIndex);
                    data = data.Replace("</ReportParametersLayout>", "");

                    //overwrite origin file
                    System.IO.File.WriteAllText(item, data);
                    removeFileList += $"{item}\r\n";
                }
            }

            VsShellUtilities.ShowMessageBox(
                this.package,
                removeFileList,
                "Remove RDLC Report Tag List",
                OLEMSGICON.OLEMSGICON_INFO,
                OLEMSGBUTTON.OLEMSGBUTTON_OK,
                OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);
        }

        /// <summary>
        /// 取得目前專案目錄
        /// </summary>
        /// <returns></returns>
        internal static string GetActiveProjectPath()
        {
            DTE dte = Package.GetGlobalService(typeof(SDTE)) as DTE;

            string projectPath = dte.ActiveDocument.Path;
            return projectPath;
        }
        #endregion





    }
}
