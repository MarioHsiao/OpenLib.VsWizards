using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio.TemplateWizard;
using System.Collections.Generic;
using System.IO;

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

        /// <summary>
        /// Defines the wizard data dictionary key for the custom project
        /// directory.
        /// </summary>
        private const string WizardDataKeyCustomProjectDir = "customProjectDir";

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
                this.MoveProject(
                    this.SolutionRoot,
                    this.WizardData,
                    this.TemplatePath,
                    this.DefaultDestinationPath,
                    this.ProjectName);
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
            base.RunStarted(automationObject,
                            replacementsDictionary,
                            runKind,
                            customParams);

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
        /// Moves a Visual Studio project to a solution folder using the
        /// specified wizard data and to an overriden physical folder on the
        /// file system.
        /// </summary>
        /// <param name="solutionRoot">Directory information for the Visual
        /// Studio solution root.</param>
        /// <param name="wizardData">A dictionary containing data for the
        /// template wizard.</param>
        /// <param name="templatePath">The path to the Visual Studio project
        /// template.</param>
        /// <param name="projectPath">The path of the Visual Studio project
        /// that was created.</param>
        /// <param name="projectName">The name of the Visual Studio project
        /// that was created.</param>
        private void MoveProject(DirectoryInfo solutionRoot,
                                 Dictionary<string, string> wizardData,
                                 string templatePath,
                                 string projectPath,
                                 string projectName)
        {
            var dte = this.VsUtils.GetActiveInstance();
            var solution = dte.Solution;

            DirectoryInfo projectRoot = new DirectoryInfo(
                this.GetNewProjectPath(solutionRoot.FullName, wizardData[WizardDataKeyCustomProjectDir], projectName));

            if (!projectRoot.Exists)
            {
                Project project = this.VsUtils.FindProject(projectName);
                Project folder = this.VsUtils.FindSolutionFolder(wizardData[WizardDataKeySolutionFolder]);

                solution.Remove(project);

                if (folder != null)
                {
                    SolutionFolder solutionFolder = this.VsUtils.ConvertToSolutionFolder(folder);
                    
                    solutionFolder.AddFromTemplate(
                        templatePath,
                        projectRoot.FullName,
                        projectName);
                    
                    this.IoUtils.DeleteDirectory(projectPath);
                }
            }
        }

        /// <summary>
        /// Gets the absolute path for a Visual Studio project using the
        /// specified solution path, path to the current location of
        /// the project, and a custom relative directory path to a new
        /// location for the project based on solution path root.
        /// </summary>
        /// <param name="solutionPath">The absolute path of the Visual Studio
        /// solution directory.</param>
        /// <param name="customProjectDir">The relative path of the optional
        /// custom project directory.</param>
        /// <param name="projectName">The sbsolute current path to the Visual
        /// Studio project.</param>
        /// <returns>The new absolute path to the project.</returns>
        private string GetNewProjectPath(string solutionPath,
                                         string customProjectDir,
                                         string projectName)
        {
            return Path.Combine(solutionPath, customProjectDir, projectName);
        }
    }
}
