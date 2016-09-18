//------------------------------------------------------------------------------
// <copyright file="CollapseEverythingElse.cs" company="Company">
//     Copyright (c) Company.  All rights reserved.
// </copyright>
//------------------------------------------------------------------------------

using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel.Design;
using System.Globalization;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.OLE.Interop;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.VisualStudio.TextManager.Interop;

namespace CollapseEverythingElse
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
                var menuCommandID = new CommandID(CommandSet, CommandId);
                var menuItem = new MenuCommand(this.MenuItemCallback, menuCommandID);
                commandService.AddCommand(menuItem);
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
        private void MenuItemCallback(object sender, EventArgs args)
        {
            IVsUIShell shell = this.ServiceProvider.GetService<SVsUIShell, IVsUIShell>();
            foreach (var windowFrame in shell.GetDocumentWindows())
            {
                try
                {
                    object docView;
                    windowFrame.GetProperty((int)__VSFPROPID.VSFPROPID_DocView, out docView);
                    var codeWindow = docView as IVsCodeWindow;

                    IVsTextView textView;
                    codeWindow.GetLastActiveView(out textView);

                    int line, column;
                    textView.GetCaretPos(out line, out column);

                    // See here for total bullshit: https://msdn.microsoft.com/en-us/library/cc826040.aspx
                    // This contains 'info' about where to find these values, but doesn't actually tell you where to look for the constants.
                    //object commandDispatcherObj = this.package.QueryService(VSConstants.SID_SUIHostCommandDispatcher);
                    //IOleCommandTarget commandDispatcher = commandDispatcherObj as IOleCommandTarget;
                    IOleCommandTarget commandDispatcher = this.ServiceProvider.GetService<SUIHostCommandDispatcher, IOleCommandTarget>();

                    Guid commandSet = VSConstants.CMDSETID.StandardCommandSet2010_guid;
                    uint commandId = (uint)VSConstants.VSStd2010CmdID.OUTLN_COLLAPSE_ALL;
                    var result = commandDispatcher.Exec(commandSet, commandId, 0, IntPtr.Zero, IntPtr.Zero);

                    textView.SetCaretPos(line, column);
                }
                catch (Exception e)
                {
                    // Show a message box to prove we were here
                    VsShellUtilities.ShowMessageBox(
                        this.ServiceProvider,
                        e.ToString(),
                        "Error",
                        OLEMSGICON.OLEMSGICON_INFO,
                        OLEMSGBUTTON.OLEMSGBUTTON_OK,
                        OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);
                }
            }
        }
    }

    public class WindowFrameEnumerable : IEnumerable<IVsWindowFrame>
    {
        private IEnumWindowFrames enumWindowFrames;

        public WindowFrameEnumerable(IVsUIShell shell)
        {
            shell.GetDocumentWindowEnum(out this.enumWindowFrames);
        }

        public WindowFrameEnumerable(IEnumWindowFrames enumWindowFrames)
        {
            this.enumWindowFrames = enumWindowFrames;
        }

        public IEnumerator<IVsWindowFrame> GetEnumerator()
        {
            return new WindowFrameEnumerator(this.enumWindowFrames);
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return this.GetEnumerator();
        }
    }

    public class WindowFrameEnumerator : IEnumerator<IVsWindowFrame>
    {
        public const int BufferSize = 5;

        private IEnumWindowFrames enumWindowFrames;
        private IVsWindowFrame[] buffer;
        private int windowCount = 0;
        private int position = 0;

        public WindowFrameEnumerator(IEnumWindowFrames enumWindowFrames)
        {
            this.enumWindowFrames = enumWindowFrames;
            this.buffer = new IVsWindowFrame[BufferSize];
        }

        public IVsWindowFrame Current
        {
            get
            {
                return this.buffer[position];
            }
        }

        object IEnumerator.Current
        {
            get
            {
                return this.Current;
            }
        }

        public bool MoveNext()
        {
            position++;
            if (position >= windowCount)
            {
                uint framesRetrieved;
                this.enumWindowFrames.Next(BufferSize, this.buffer, out framesRetrieved);
                this.windowCount = (int)framesRetrieved;
                position = 0;

                if (this.windowCount == 0)
                {
                    return false;
                }
            }

            return true;
        }

        public void Reset()
        {
            this.position = 0;
            this.windowCount = 0;
            this.enumWindowFrames.Reset();
        }

        #region IDisposable Support
        private bool disposedValue = false; // To detect redundant calls

        protected virtual void Dispose(bool disposing)
        {
            if (!disposedValue)
            {
                if (disposing)
                {

                }

                disposedValue = true;
            }
        }

        public void Dispose()
        {
            Dispose(true);
        }
        #endregion
    }

    public static class EnumWindowFramesExtensions
    {
        public static WindowFrameEnumerator GetEnumerator(this IEnumWindowFrames self)
        {
            return new WindowFrameEnumerator(self);
        }
    }

    public static class IVsUIShellExtensions
    {
        public static WindowFrameEnumerable GetDocumentWindows(this IVsUIShell shell)
        {
            return new WindowFrameEnumerable(shell);
        }
    }

    public static class IServiceProviderExtensions
    {
        public static TIService GetService<TService, TIService>(this System.IServiceProvider serviceProvider)
        {
            return (TIService)serviceProvider.GetService(typeof(TService));
        }
    }
}
