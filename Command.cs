using EnvDTE80;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System;
using System.ComponentModel.Design;
using System.Globalization;
using System.Threading;
using System.Threading.Tasks;
using Task = System.Threading.Tasks.Task;
using System.IO;
using EnvDTE;
using Microsoft.Build.Evaluation;
//using Microsoft.Build.Locator;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using Project = Microsoft.Build.Evaluation.Project;
using System.Text.RegularExpressions;

namespace SetUpForDebug
{
    /// <summary>
    /// Command handler
    /// </summary>
    internal sealed class Command
    {
        /// <summary>
        /// Command ID.
        /// </summary>
        public const int CommandId = 0x0100;

        /// <summary>
        /// Command menu group (command set GUID).
        /// </summary>
        public static readonly Guid CommandSet = new Guid("f40b03a5-f476-4645-93b1-10d3b0493b37");

        /// <summary>
        /// VS Package that provides this command, not null.
        /// </summary>
        private readonly AsyncPackage package;

        /// <summary>
        /// Initializes a new instance of the <see cref="Command"/> class.
        /// Adds our command handlers for menu (commands must exist in the command table file)
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        /// <param name="commandService">Command service to add command to, not null.</param>
        private Command(AsyncPackage package, OleMenuCommandService commandService)
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
        public static Command Instance
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

        /// <summary>
        /// Initializes the singleton instance of the command.
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        public static async Task InitializeAsync(AsyncPackage package)
        {
            // Switch to the main thread - the call to AddCommand in Command's constructor requires
            // the UI thread.
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(package.DisposalToken);

            OleMenuCommandService commandService = await package.GetServiceAsync(typeof(IMenuCommandService)) as OleMenuCommandService;
            Instance = new Command(package, commandService);
        }

        /// <summary>
        /// This function is the callback used to execute the command when the menu item is clicked.
        /// See the constructor to see how the menu item is associated with this function using
        /// OleMenuCommandService service and MenuCommand class.
        /// </summary>
        /// <param name="sender">Event sender.</param>
        /// <param name="e">Event args.</param>
        private async void Execute(object sender, EventArgs e)
        {

            try
            {
                await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();

                var dte = await ServiceProvider.GetServiceAsync(typeof(DTE)) as DTE2;
                if (dte == null)
                {
                    return;
                }

                var activeDoc = dte.ActiveDocument;
                if (activeDoc == null || !activeDoc.Name.EndsWith(".csproj"))
                {
                    System.Windows.Forms.MessageBox.Show("Please run in the project properties");
                    return;
                }

                var projPath = activeDoc.FullName;
                var projDirectory = Path.GetDirectoryName(projPath);
                var solutionPath = dte.Solution.FullName;

                var proj = new Project(projPath);
                string updatedUrl = GenerateDebugUrl(projPath);

                // Create virtual direcoty (if needed)

                // Set the start URL to virtual directory with debug.auburn.edu (add no start url feature/setting?)
                SetStartUrl(proj, updatedUrl, projPath);

                // Add debug binding in applcationhost file


                string message = $"Setting up for Debug..... doc is {activeDoc.Name}";
                string title = "Command";

                // Show a message box to prove we were here
                VsShellUtilities.ShowMessageBox(
                    this.package,
                    message,
                    title,
                    OLEMSGICON.OLEMSGICON_INFO,
                    OLEMSGBUTTON.OLEMSGBUTTON_OK,
                    OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);}
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
            }
        }

        // takes project url (local host) and makes the debug url (that will be the start url)
        private string GenerateDebugUrl(string projPath)
        {
            var code = XDocument.Load(projPath);
            var whatisns = code.Root.GetDefaultNamespace();
            var found = code.Descendants(whatisns + "IISUrl").Select(de => de.Value);
            string iisUrlInnerHtml = string.Join(Environment.NewLine, found);
            return Regex.Replace(iisUrlInnerHtml, @"http://localhost", "http://debug.auburn.edu");
        }

        // start url should be: http://localhost:8080/whatver-the-project-is with "local host" as "debug.auburn.edu"
        private void SetStartUrl(Project proj, string startUrl, string projPath)
        {
            // Load the .csproj file
            var code = XDocument.Load(projPath);
            XDocument userFile = XDocument.Load(projPath + ".user");

            XElement projExtension = new XElement("ProjectExtensions",
                new XElement("VisualStudio",
                    new XElement("FlavorProperties", new XAttribute("GUID", "placeholder"),
                        new XElement("WebProjectProperties",
                            new XElement("StartPageUrl"),
                            new XElement("StartAction", "URL"),
                            new XElement("AspNetDebugging", "True"),
                            new XElement("SilverlightDebugging", "False"),
                            new XElement("NativeDebugging", "False"),
                            new XElement("SQLDebugging", "False"),
                            new XElement("ExternalProgram"),
                            new XElement("StartExternalURL", startUrl),
                            new XElement("StartCmdLineArguments"),
                            new XElement("StartWorkingDirectory"),
                            new XElement("EnableENC", "True"),
                            new XElement("AlwaysStartWebServerOnDebug", "False")
                        )

                    )
                )
            );
            Console.WriteLine("done");

        }

        //private void AddDebugBinding(string solutionPath, string projDirectory)
        //{
        //    var solutionDirectory = Path.GetDirectoryName(solutionPath);
        //    var applicationHostPath = Path.Combine(solutionDirectory, ".vs", "config", "applicationhost.config");
        //    var applicationHost = XDocument.Load(applicationHostPath);
        //    var site = applicationHost.Descendants("site").FirstOrDefault();
        //    var bindings = site.Descendants("bindings").FirstOrDefault();
        //    var binding = new XElement("binding");
        //    binding.SetAttributeValue("protocol", "http");
        //    binding.SetAttributeValue("bindingInformation", "*:80:debug.auburn.edu");
        //    bindings.Add(binding);
        //    applicationHost.Save(applicationHostPath);
        //}
    }
}
