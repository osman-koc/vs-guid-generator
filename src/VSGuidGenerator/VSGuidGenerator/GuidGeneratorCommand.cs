using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System;
using System.ComponentModel.Design;
using Task = System.Threading.Tasks.Task;

namespace VSGuidGenerator
{
    /// <summary>
    /// Command handler
    /// </summary>
    internal sealed class GuidGeneratorCommand
    {
        /// <summary>
        /// Command ID.
        /// </summary>
        public const int CommandId = 0x0100;

        /// <summary>
        /// Command menu group (command set GUID).
        /// </summary>
        public static readonly Guid CommandSet = new Guid("86c0f708-9d77-44a7-85a1-da159a9852a0");

        /// <summary>
        /// VS Package that provides this command, not null.
        /// </summary>
        private readonly AsyncPackage package;

        /// <summary>
        /// Initializes a new instance of the <see cref="GuidGeneratorCommand"/> class.
        /// Adds our command handlers for menu (commands must exist in the command table file)
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        /// <param name="commandService">Command service to add command to, not null.</param>
        private GuidGeneratorCommand(AsyncPackage package, OleMenuCommandService commandService)
        {
            this.package = package ?? throw new ArgumentNullException(nameof(package));
            commandService = commandService ?? throw new ArgumentNullException(nameof(commandService));

            var menuCommandID = new CommandID(CommandSet, CommandId);
            var menuItem = new MenuCommand(this.Execute, menuCommandID);
            commandService.AddCommand(menuItem);
        }

        /// <summary>
        /// Gets the instance of the command.
        /// </summary>
        public static GuidGeneratorCommand Instance
        {
            get;
            private set;
        }

        /// <summary>
        /// Gets the service provider from the owner package.
        /// </summary>
        private IAsyncServiceProvider ServiceProvider
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
        public static async Task InitializeAsync(AsyncPackage package)
        {
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(package.DisposalToken);

            var commandService = await package.GetServiceAsync(typeof(IMenuCommandService)) as OleMenuCommandService;
            Instance = new GuidGeneratorCommand(package, commandService);
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

            var dte2 = (DTE2)Package.GetGlobalService(typeof(SDTE));
            if (dte2 == null) return;

            var activeDocument = dte2.ActiveDocument;
            if (activeDocument == null) return;

            var textDocument = activeDocument.Object() as TextDocument;
            if (textDocument == null) return;

            var sel = textDocument.Selection as TextSelection;
            if (sel == null) return;

            WriteGuidString(textDocument, sel);
        }

        /// <summary>
        /// This method generates a new GUID string and inserts it at the current cursor position or selected text. 
        /// If multiple lines are selected, it generates and inserts GUID strings for each line in the selected range, 
        /// ensuring each line receives a unique GUID. The method provides the ability to efficiently insert GUIDs 
        /// for both single-line and multi-line scenarios within a text document.
        /// </summary>
        /// <param name="sel"></param>
        private void WriteGuidString(TextDocument textDoc, TextSelection sel)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            if (sel == null) return;

            var startLine = sel.TopPoint.Line;
            var endLine = sel.BottomPoint.Line;

            //var lineStartColumn = sel.AnchorPoint.LineCharOffset - 1;
            var lineEndColumn = sel.ActivePoint.LineCharOffset;

            for (int i = startLine; i <= endLine; i++)
            {
                var editPoint = textDoc.CreateEditPoint(textDoc.StartPoint);
                var lineText = editPoint.GetLines(i, i + 1);
                var newGuid = Guid.NewGuid().ToString().ToUpper();

                var startIndex = lineEndColumn - 1;
                if (startIndex > lineText.Length)
                {
                    startIndex = lineText.Length;
                }
                var updatedLineText = lineText.Insert(startIndex, newGuid);

                sel.GotoLine(i);
                sel.StartOfLine(vsStartOfLineOptions.vsStartOfLineOptionsFirstColumn, false);
                sel.Delete(lineText.Length);
                sel.Insert(updatedLineText);
            }
        }

    }
}
