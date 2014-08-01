using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio.TemplateWizard;
using OpenLib.Extensions;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace OpenLib.VsWizards
{
    /// <summary>
    /// The <c>FolderWizard</c> type provides an implementation of the
    /// <c>IWizard</c> interface that provides a custom Visual Studio project
    /// template wizard for overriding the default behavior of where physical
    /// project folders are created.
    /// </summary>
    /// <remarks>
    /// This class inherits from AbstractWizard, which provides common
    /// Wizard properties and methods.
    /// </remarks>
    public class FolderWizard : AbstractWizard, IWizard
    {
        //---------------------------------------------------------------------
        // Fields
        //---------------------------------------------------------------------

        /// <summary>
        /// Defines the wizard data dictionary key for the project template.
        /// </summary>
        private const string WizardDataKeyTemplate = "template";

        /// <summary>
        /// Defines the wizard data dictionary key for the solution folder.
        /// </summary>
        private const string WizardDataKeySolutionFolder = "solutionFolder";

        //---------------------------------------------------------------------
        // Constructors
        //---------------------------------------------------------------------

        /// <summary>
        /// Creates a new instance of the <c>FolderWizard</c> class.
        /// </summary>
        public FolderWizard() : base()
        {
        }

        //---------------------------------------------------------------------
        // Abstract Implementation Methods
        //---------------------------------------------------------------------

        /// <summary>
        /// Validates the creation of a Visual Studio project using a project
        /// template from this Wizard class implementation.
        /// </summary>
        /// <returns>A value indicating if the project template is valid.</returns>
        protected override bool Validate()
        {
            return this.WizardData != null &&
                    this.WizardData.ContainsKey(WizardDataKeyTemplate) &&
                    this.WizardData.ContainsKey(WizardDataKeySolutionFolder);
        }

        //---------------------------------------------------------------------
        // IWizard Implementation
        //---------------------------------------------------------------------

        /// <summary>
        /// Executes before each file that is being added to a Visual Studio
        /// project is opened.
        /// </summary>
        /// <param name="projectItem">A reference to the project item being
        /// added to the Visual Studio project.</param>
        public void BeforeOpeningFile(ProjectItem projectItem)
        {
        }

        /// <summary>
        /// Executes when a Visual Studio project is being created.
        /// </summary>
        /// <param name="project">A reference to the Visual Studio project
        /// being created.</param>
        public void ProjectFinishedGenerating(Project project)
        {
        }

        /// <summary>
        /// Executes when a Visual Studio project item is created.
        /// </summary>
        /// <param name="projectItem">A reference to the project item created
        /// in the Visual Studio project.</param>
        public void ProjectItemFinishedGenerating(ProjectItem projectItem)
        {
        }

        /// <summary>
        /// Executes when the creation of a Visual Studio project is complete.
        /// </summary>
        public void RunFinished()
        {
            if (this.IsValid)
            {
                this.AddProject(
                    this.SolutionRoot,
                    this.WizardData,
                    this.TemplatePath,
                    this.DefaultDestinationPath,
                    this.NameSpace);
            }
        }

        /// <summary>
        /// Gets a value indicating if a project item should be added to a
        /// Visual Studio project.
        /// </summary>
        /// <param name="filePath">The absolute path to the project item being
        /// added.</param>
        /// <returns>A value indicating if a project item should be added.</returns>
        public bool ShouldAddProjectItem(string filePath)
        {
            return true;
        }

        /// <summary>
        /// Executes when the creation of a Visual Studio project is started.
        /// </summary>
        /// <param name="automationObject">A reference to the automation object
        /// used by the Visual Studio project template wizard.</param>
        /// <param name="replacementsDictionary">A reference to a dictionary of
        /// template variables that can be modified to provide customization.</param>
        /// <param name="runKind">Defines the type of the template the template
        /// wizard is creating.</param>
        /// <param name="customParams">An array of custom parameters passed to
        /// the template wizard.</param>
        public new void RunStarted(object automationObject,
                                   Dictionary<string, string> replacementsDictionary,
                                   WizardRunKind runKind,
                                   object[] customParams)
        {
            base.RunStarted(automationObject, replacementsDictionary, runKind, customParams);

            if (this.IsValid)
            {
                this.TemplatePath = Path.Combine(
                    Path.GetDirectoryName(this.TemplatePath),
                    this.WizardData[WizardDataKeyTemplate]);
            }
        }

        //---------------------------------------------------------------------
        // Other Methods
        //---------------------------------------------------------------------

        /// <summary>
        /// Adds a Visual Studio project to a solution folder using the
        /// specified wizard data and to an overriden physical folder on the
        /// file system.
        /// </summary>
        /// <param name="solutionRoot">Directory information for the Visual
        /// Studio solution root.</param>
        /// <param name="wizardData">A dictionary containing data for the
        /// template wizard.</param>
        /// <param name="templatePath">The path to the Visual Studio project
        /// template.</param>
        /// <param name="defaultDestinationPath">The default destination path
        /// of the Visual Studio project to be created.</param>
        /// <param name="nameSpace">The namespace for the Visual Studio project
        /// to be created.</param>
        private void AddProject(DirectoryInfo solutionRoot,
                                Dictionary<string, string> wizardData,
                                string templatePath,
                                string defaultDestinationPath,
                                string nameSpace)
        {
            var dte = this.VsUtils.GetActiveInstance();
            var solution = dte.Solution;

            DirectoryInfo destination = new DirectoryInfo(
                this.StripSolutionPath(defaultDestinationPath, solutionRoot.Name));

            if (!destination.Exists)
            {
                Project project = this.VsUtils.FindProject(nameSpace);
                Project folder = this.VsUtils.FindSolutionFolder(wizardData[WizardDataKeySolutionFolder]);

                solution.Remove(project);

                if (folder != null)
                {
                    SolutionFolder solutionFolder = this.VsUtils.ConvertToSolutionFolder(folder);

                    solutionFolder.AddFromTemplate(
                        templatePath,
                        destination.FullName,
                        nameSpace);
                }
            }
        }

        /// <summary>
        /// Strips an extra directory out of the specified solution path to
        /// override the default Visual Studio project folder's location on the
        /// file system.
        /// </summary>
        /// <param name="currentPath">The current path to the Visual Studio
        /// solution.</param>
        /// <param name="solutionDir">The name of the Visual Studio solution
        /// directory.</param>
        /// <returns>The updated solution path with the extra directory stripped
        /// out.</returns>
        private string StripSolutionPath(string currentPath, string solutionDir)
        {
            string path = currentPath;
            IEnumerable<int> enumerable = path.IndexOfAll(solutionDir);

            if (enumerable.Count() == 3)
            {
                int index = enumerable.ElementAt(1);
                path = path.Remove(index, solutionDir.Length);
            }

            return path;
        }
    }
}
