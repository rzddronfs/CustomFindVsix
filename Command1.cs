﻿using System;
using System.Collections.Generic;
using System.ComponentModel.Design;
using System.Globalization;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Task = System.Threading.Tasks.Task;
using Dispatcher = System.Windows.Threading.Dispatcher;


namespace CustomFindVsix
{
  /// <summary>
  /// Command handler
  /// </summary>
  internal sealed class FindEntriesCmd
  {
    /// <summary>
    /// Command ID.
    /// </summary>
    public const int ProjectScopeId = 256;
    public const int DocumentScopeId = 257;

    /// <summary>
    /// Command menu group (command set GUID).
    /// </summary>
    public static readonly Guid CommandSet = new Guid("a3490c9f-4115-4d25-b99c-17c6951f68db");

    /// <summary>
    /// VS Package that provides this command, not null.
    /// </summary>
    private readonly AsyncPackage package;

    private readonly int[] Cmds = new[] { ProjectScopeId, DocumentScopeId };
    private readonly Dictionary<int, Runnable> ExecOptions;
    
    private delegate void Runnable(object sender, EventArgs e);

    /// <summary>
    /// Initializes a new instance of the <see cref="FindEntriesCmd"/> class.
    /// Adds our command handlers for menu (commands must exist in the command table file)
    /// </summary>
    /// <param name="package">Owner package, not null.</param>
    /// <param name="commandService">Command service to add command to, not null.</param>
    private FindEntriesCmd(AsyncPackage package, OleMenuCommandService commandService)
    {
      if (package == null)
        throw new ArgumentNullException();

      if (commandService == null)
        throw new ArgumentNullException();

      this.package = package;

      ExecOptions = new Dictionary<int, Runnable>
      {
        {ProjectScopeId, ExecuteOnProject},
        {DocumentScopeId, ExecuteOnDocument},
      };

      System.Diagnostics.Debug.Assert(ExecOptions.Count == Cmds.Length);

      foreach (var cmdId in Cmds)
      {
        var menuCommandID = new CommandID(CommandSet, cmdId);
        MenuCommand menuItem = new MenuCommand(
          (s, e) => 
          {
            Dispatcher.CurrentDispatcher.VerifyAccess();
            Execute(cmdId, s, e); 
          }, 
          menuCommandID
        );
        commandService.AddCommand(menuItem);
      }
    }

    /// <summary>
    /// Gets the instance of the command.
    /// </summary>
    public static FindEntriesCmd Instance
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
      // Switch to the main thread - the call to AddCommand in FindEntriesCmd's constructor requires
      // the UI thread.
      await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(package.DisposalToken);

      OleMenuCommandService commandService = await package.GetServiceAsync((typeof(IMenuCommandService))) as OleMenuCommandService;
      Instance = new FindEntriesCmd(package, commandService);
    }

    /// <summary>
    /// This function is the callback used to execute the command when the menu item is clicked.
    /// See the constructor to see how the menu item is associated with this function using
    /// OleMenuCommandService service and MenuCommand class.
    /// </summary>
    /// <param name="sender">Event sender.</param>
    /// <param name="e">Event args.</param>
    private void Execute( int cmdId, object sender, EventArgs e)
    {
      ThreadHelper.ThrowIfNotOnUIThread();
      var dte = Package.GetGlobalService(typeof(SDTE)) as EnvDTE._DTE;
      
      if ((dte.ActiveDocument.Selection as EnvDTE.TextSelection).Text.Length == 0)
        return;

      ExecOptions[cmdId](sender, e);
    }

    private void ExecuteOnProject(object sender, EventArgs e)
    {
      ThreadHelper.ThrowIfNotOnUIThread();
      var dte = Package.GetGlobalService(typeof(SDTE)) as EnvDTE._DTE;
      var seeker = dte.Find;

      seeker.Action = EnvDTE.vsFindAction.vsFindActionFindAll;
      seeker.FindWhat = (dte.ActiveDocument.Selection as EnvDTE.TextSelection).Text;
      seeker.ResultsLocation = EnvDTE.vsFindResultsLocation.vsFindResults2;

      seeker.Backwards = false;
      seeker.MatchCase = false;
      seeker.MatchInHiddenText = true;
      seeker.MatchWholeWord = false;
      seeker.SearchSubfolders = true;
      seeker.PatternSyntax = EnvDTE.vsFindPatternSyntax.vsFindPatternSyntaxLiteral;
      seeker.Target = EnvDTE.vsFindTarget.vsFindTargetCurrentProject;
      seeker.Execute();
    }

    private void ExecuteOnDocument(object sender, EventArgs e)
    {
      ThreadHelper.ThrowIfNotOnUIThread();
      var dte = Package.GetGlobalService(typeof(SDTE)) as EnvDTE._DTE;
      var seeker = dte.Find;

      seeker.Action = EnvDTE.vsFindAction.vsFindActionFindAll;
      seeker.FindWhat = (dte.ActiveDocument.Selection as EnvDTE.TextSelection).Text;
      seeker.ResultsLocation = EnvDTE.vsFindResultsLocation.vsFindResults2;

      seeker.Backwards = false;
      seeker.MatchCase = false;
      seeker.MatchInHiddenText = true;
      seeker.MatchWholeWord = false;
      seeker.SearchSubfolders = true;
      seeker.PatternSyntax = EnvDTE.vsFindPatternSyntax.vsFindPatternSyntaxLiteral;
      seeker.Target = EnvDTE.vsFindTarget.vsFindTargetCurrentDocument;
      seeker.Execute();
    }

  }
}
