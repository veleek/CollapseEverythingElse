//------------------------------------------------------------------------------
// <copyright file="CollapseEverythingElse.cs" company="Ben Corp.">
//     Copyright (c) Ben Corp.  All rights reserved.
// </copyright>
//------------------------------------------------------------------------------

using System;
using System.ComponentModel.Design;
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
        private readonly Package package;

        /// <summary>
        /// Initializes a new instance of the <see cref="CollapseEverythingElse"/> class.
        /// Adds our command handlers for menu (commands must exist in the command table file)
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        private CollapseEverythingElse(Package package)
        {
            if (package == null)
            {
                throw new ArgumentNullException("package");
            }

            this.package = package;

            OleMenuCommandService commandService = this.ServiceProvider.GetService(typeof(IMenuCommandService)) as OleMenuCommandService;
            if (commandService != null)
            {
                var collapseEverythingElseCommandID = new CommandID(CommandSet, CommandId);
                var collapseEverythingElseMenuItem = new MenuCommand(this.CollapseEverythingElseCallback, collapseEverythingElseCommandID);
                commandService.AddCommand(collapseEverythingElseMenuItem);
            }
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
        private System.IServiceProvider ServiceProvider
        {
            get
            {
                return this.package;
            }
        }

        /// <summary>
        /// Initializes the singleton instance of the command.
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        public static void Initialize(Package package)
        {
            Instance = new CollapseEverythingElse(package);
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
            IVsTextManager textManager = this.ServiceProvider.GetService<SVsTextManager, IVsTextManager>();
            IVsTextView activeView;
            textManager.GetActiveView(0, null, out activeView);

            CollapseAllRegionsExceptCurrent(activeView);
        }

        private void CollapseEverythingElseInAllWindows()
        {
            IVsUIShell shell = this.ServiceProvider.GetService<SVsUIShell, IVsUIShell>();
            foreach (IVsWindowFrame windowFrame in shell.GetDocumentWindows())
            {
                var codeWindow = windowFrame.GetProperty<IVsCodeWindow>(__VSFPROPID.VSFPROPID_DocView);

                IVsTextView textView;
                codeWindow.GetLastActiveView(out textView);

                CollapseAllRegionsExceptCurrent(textView);
            }
        }

        /// <summary>
        /// Collapse all regions except for the current region (and any parents) in the given TextView.
        /// </summary>
        /// <param name="textView">The text view to collapse.</param>
        private void CollapseAllRegionsExceptCurrent(IVsTextView textView)
        {
            int line, column;
            textView.GetCaretPos(out line, out column);

            IOleCommandTarget commandDispatcher = this.ServiceProvider.GetService<SUIHostCommandDispatcher, IOleCommandTarget>();

            // See the MSDN documentation here https://msdn.microsoft.com/en-us/library/cc826040.aspx
            // for pointers about where command set and command IDs are defined.  Specifically, we're
            // using Edit > Outlining > Collapse All.
            Guid commandSet = VSConstants.CMDSETID.StandardCommandSet2010_guid;
            uint commandId = (uint)VSConstants.VSStd2010CmdID.OUTLN_COLLAPSE_ALL;
            var result = commandDispatcher.Exec(commandSet, commandId, 0, IntPtr.Zero, IntPtr.Zero);

            textView.SetCaretPos(line, column);
            textView.CenterLines(line, 10);
        }
    }
}
