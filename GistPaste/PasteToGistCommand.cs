//------------------------------------------------------------------------------
// <copyright file="PasteToGistCommand.cs" company="Company">
//     Copyright (c) Company.  All rights reserved.
// </copyright>
//------------------------------------------------------------------------------

using System;
using System.ComponentModel.Design;
using Microsoft.VisualStudio.Editor;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.VisualStudio.Text.Editor;
using Microsoft.VisualStudio.TextManager.Interop;

namespace GistPaste
{
    /// <summary>
    /// Command handler
    /// </summary>
    internal sealed class PasteToGistCommand
    {
        /// <summary>
        /// Command ID.
        /// </summary>
        public const int CommandId = 0x0100;

        /// <summary>
        /// Command menu group (command set GUID).
        /// </summary>
        public static readonly Guid CommandSet = new Guid ("b0b6a288-0e8c-476d-93e8-eaf9d96c504e");

        /// <summary>
        /// VS Package that provides this command, not null.
        /// </summary>
        private readonly Package package;

        /// <summary>
        /// Initializes a new instance of the <see cref="PasteToGistCommand"/> class.
        /// Adds our command handlers for menu (commands must exist in the command table file)
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        private PasteToGistCommand (Package package)
        {
            if (package == null)
                throw new ArgumentNullException (nameof (package));

            this.package = package;

            var commandService = ServiceProvider.GetService (typeof (IMenuCommandService)) as OleMenuCommandService;

            if (commandService != null) {
                var menuCommandID = new CommandID (CommandSet, CommandId);
                var menuItem = new MenuCommand (MenuItemCallback, menuCommandID);

                commandService.AddCommand (menuItem);
            }
        }

        /// <summary>
        /// Gets the instance of the command.
        /// </summary>
        public static PasteToGistCommand Instance
        {
            get;
            private set;
        }

        /// <summary>
        /// Gets the service provider from the owner package.
        /// </summary>
        private IServiceProvider ServiceProvider
        {
            get
            {
                return package;
            }
        }

        /// <summary>
        /// Initializes the singleton instance of the command.
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        public static void Initialize (Package package)
        {
            Instance = new PasteToGistCommand (package);
        }

        /// <summary>
        /// This function is the callback used to execute the command when the menu item is clicked.
        /// See the constructor to see how the menu item is associated with this function using
        /// OleMenuCommandService service and MenuCommand class.
        /// </summary>
        /// <param name="sender">Event sender.</param>
        /// <param name="e">Event args.</param>
        private void MenuItemCallback (object sender, EventArgs e)
        {
            string message = GetTextToPaste ();
            string title = "Paste To Gist";

            // Show a message box to prove we were here
            VsShellUtilities.ShowMessageBox (
                this.ServiceProvider,
                message,
                title,
                OLEMSGICON.OLEMSGICON_INFO,
                OLEMSGBUTTON.OLEMSGBUTTON_OK,
                OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);
        }

        private string GetTextToPaste ()
        {
            // Get the current text manager
            var txtMgr = (IVsTextManager)ServiceProvider.GetService (typeof (SVsTextManager));

            if (txtMgr == null) {
                Console.WriteLine ("No text view is currently open");
                return null;
            }

            // Get the current text view
            IVsTextView currentTextView;
            txtMgr.GetActiveView (1, null, out currentTextView);

            // Convert to an IVsUserData
            var userData = currentTextView as IVsUserData;

            if (userData == null) {
                Console.WriteLine ("No text view is currently open");
                return null;
            }

            // Try to get access to the editor's view 
            object holder;
            Guid guidViewHost = DefGuidList.guidIWpfTextViewHost;

            userData.GetData (ref guidViewHost, out holder);
            var viewHost = holder as IWpfTextViewHost;

            if (viewHost == null) {
                Console.WriteLine ("No text view is currently open");
                return null;
            }

            if (viewHost.TextView.Selection.IsEmpty) {
                // Return entire text for the editor
                return viewHost.TextView.TextSnapshot.GetText ();
            } else {
                // Return the currently selected text
                return viewHost.TextView.Selection.SelectedSpans[0].GetText ();
            }
        }
    }
}
