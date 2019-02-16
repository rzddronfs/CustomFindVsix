//------------------------------------------------------------------------------
// <copyright file="FindEntriesCmds.cs" company="Dronfs">
//     Copyright (c) Dronfs.  All rights reserved.
// </copyright>
//------------------------------------------------------------------------------

using System;
using System.Collections.Generic;
using System.ComponentModel.Design;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;


namespace CustomFindVsix
{
  /// <summary>
  /// Command handler
  /// </summary>
  internal sealed class FindEntriesCmds
  {
    /// <summary>
    /// Command ID.
    /// </summary>
    public const int ProjectScopeId = 256;
    public const int DocumentScopeId = 257;

    /// <summary>
    /// Command menu group (command set GUID).
    /// </summary>
    public static readonly Guid CommandSet = new Guid("9965b1e1-829a-4194-bbf6-a535f6a2e51f");

    /// <summary>
    /// VS Package that provides this command, not null.
    /// </summary>
    private readonly Package package;

    private readonly int[] Cmds = new[] { ProjectScopeId, DocumentScopeId };
    private readonly Dictionary<int, Runnable> ExecOptions;
    
    private delegate void Runnable(object sender, EventArgs e);

    /// <summary>
    /// Initializes a new instance of the <see cref="FindEntriesCmds"/> class.
    /// Adds our command handlers for menu (commands must exist in the command table file)
    /// </summary>
    /// <param name="package">Owner package, not null.</param>
    private FindEntriesCmds(Package package, OleMenuCommandService commandService)
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
    public static FindEntriesCmds Instance
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
        return this.package;
      }
    }

    /// <summary>
    /// Initializes the singleton instance of the command.
    /// </summary>
    /// <param name="package">Owner package, not null.</param>
    public static void Initialize(Package package)
    {
      OleMenuCommandService commandService = 
        (package as IServiceProvider).GetService(typeof(IMenuCommandService)) as OleMenuCommandService;
      
      Instance = new FindEntriesCmds(package, commandService);
    }

    /// <summary>
    /// This function is the callback used to execute the command when the menu item is clicked.
    /// See the constructor to see how the menu item is associated with this function using
    /// OleMenuCommandService service and MenuCommand class.
    /// </summary>
    /// <param name="cmdId">Command id.</param>
    /// <param name="sender">Event sender.</param>
    /// <param name="e">Event args.</param>
    private void Execute(int cmdId, object sender, EventArgs e)
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
