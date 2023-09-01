//------------------------------------------------------------------------------
// <copyright file="CollapseEverythingElse.cs" company="Ben Corp.">
//     Copyright (c) Ben Corp.  All rights reserved.
// </copyright>
//------------------------------------------------------------------------------

using System;
using System.ComponentModel.Design;
using System.Threading.Tasks;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.OLE.Interop;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.VisualStudio.TextManager.Interop;

namespace Ben.VisualStudio
{
    /// <summary>
    /// Command handler
    /// </summary>
    internal sealed class CollapseEverythingElse
    {
        /// <summary>
        /// Command ID.
        /// </summary>
        public const int CommandId = 0x0100;

        /// <summary>
        /// Command menu group (command set GUID).
        /// </summary>
        public static readonly Guid CommandSet = new Guid("4db49a2b-c9c5-4cef-92e6-03740ddb87a7");

        /// <summary>
        /// VS Package that provides this command, not null.
        /// </summary>
        private readonly AsyncPackage package;

		/// <summary>
		/// Initializes a new instance of the <see cref="CollapseEverythingElse"/> class.
		/// Adds our command handlers for menu (commands must exist in the command table file)
		/// </summary>
		/// <param name="package">Owner package, not null.</param>
		/// <param name="commandService">Command service to add command to, not null.</param>
		private CollapseEverythingElse(AsyncPackage package, OleMenuCommandService commandService)
        {
	        this.package = package ?? throw new ArgumentNullException(nameof(package));
	        commandService = commandService ?? throw new ArgumentNullException(nameof(commandService));
            
            var collapseEverythingElseCommandID = new CommandID(CommandSet, CommandId);
            var collapseEverythingElseMenuItem = new MenuCommand(this.CollapseEverythingElseCallback, collapseEverythingElseCommandID);

            commandService.AddCommand(collapseEverythingElseMenuItem);
        }

        /// <summary>
        /// Gets the instance of the command.
        /// </summary>
        public static CollapseEverythingElse Instance
        {
            get;
            private set;
        }

        /// <summary>
        /// Gets the service provider from the owner package.
        /// </summary>
        private Microsoft.VisualStudio.Shell.IAsyncServiceProvider ServiceProvider => this.package;

        private IVsTextManager textManager;
        private IVsUIShell shell;
        private IOleCommandTarget commandDispatcher;

		/// <summary>
		/// Initializes the singleton instance of the command.
		/// </summary>
		/// <param name="package">Owner package, not null.</param>
		public static async Task InitializeAsync(AsyncPackage package)
        {
            // Switch to the main thread - the call to AddCommand in the constructor requires the UI thread.
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(package.DisposalToken);

            
            OleMenuCommandService commandService = await package.GetServiceAsync<IMenuCommandService, OleMenuCommandService>();
            Instance = new CollapseEverythingElse(package, commandService)
            {
	            textManager = await package.GetServiceAsync<SVsTextManager, IVsTextManager>(),
	            shell = await package.GetServiceAsync<SVsUIShell, IVsUIShell>(),
	            commandDispatcher = await package.GetServiceAsync<SUIHostCommandDispatcher, IOleCommandTarget>()
            };
        }

        /// <summary>
        /// This function is the callback used to execute the command when the menu item is clicked.
        /// See the constructor to see how the menu item is associated with this function using
        /// OleMenuCommandService service and MenuCommand class.
        /// </summary>
        /// <param name="sender">Event sender.</param>
        /// <param name="e">Event args.</param>
        private void CollapseEverythingElseCallback(object sender, EventArgs args)
        {
            CollapseEverythingElseInActiveWindow();
        }

        private void CollapseEverythingElseInActiveWindow()
        {
	        textManager.GetActiveView(0, null, out IVsTextView activeView);
            CollapseAllRegionsExceptCurrent(activeView);
        }

        private void CollapseEverythingElseInAllWindows()
        {
            foreach (IVsWindowFrame windowFrame in shell.GetDocumentWindows())
            {
                var codeWindow = windowFrame.GetProperty<IVsCodeWindow>(__VSFPROPID.VSFPROPID_DocView);

                codeWindow.GetLastActiveView(out IVsTextView textView);
                CollapseAllRegionsExceptCurrent(textView);
            }
        }

        /// <summary>
        /// Collapse all regions except for the current region (and any parents) in the given TextView.
        /// </summary>
        /// <param name="textView">The text view to collapse.</param>
        private void CollapseAllRegionsExceptCurrent(IVsTextView textView)
        {
			ThreadHelper.ThrowIfNotOnUIThread();
			textView.GetCaretPos(out int line, out int column);

            // See the MSDN documentation here https://msdn.microsoft.com/en-us/library/cc826040.aspx
            // for pointers about where command set and command IDs are defined.  Specifically, we're
            // using Edit > Outlining > Collapse All.
            Guid commandSet = VSConstants.CMDSETID.StandardCommandSet2010_guid;
            const uint commandId = (uint)VSConstants.VSStd2010CmdID.OUTLN_COLLAPSE_ALL;
            commandDispatcher.Exec(commandSet, commandId, 0, IntPtr.Zero, IntPtr.Zero);

            textView.SetCaretPos(line, column);
            textView.CenterLines(line, 10);
        }
    }
}
