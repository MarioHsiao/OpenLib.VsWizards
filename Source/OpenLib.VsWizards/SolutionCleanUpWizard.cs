using EnvDTE;
using Microsoft.VisualStudio.TemplateWizard;
using System.Collections.Generic;
using System.IO;

namespace OpenLib.VsWizards
{
    /// <summary>
    /// The <c>SolutionCleanUpWizard</c> type provides an implementation of the
    /// <c>IWizard</c> interface that provides a custom Visual Studio project
    /// template wizard for cleaning up Visual Studio solution and project
    /// directories after a solution is created.
    /// </summary>
    /// <remarks>
    /// This class inherits from AbstractWizard, which provides common
    /// Wizard properties and methods.
    /// </remarks>
    public class SolutionCleanUpWizard : AbstractWizard, IWizard
    {
        //---------------------------------------------------------------------
        // Constructors
        //---------------------------------------------------------------------

        /// <summary>
        /// Creates a new instance of the <c>SolutionCleanUpWizard</c> class.
        /// </summary>
        public SolutionCleanUpWizard() : base()
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
            return true;
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
            this.CleanSolution(this.SolutionRoot);
        }

        /// <summary>
        /// Gets a value indicating if a project item should be added to a
        /// Visual Studio project.
        /// </summary>
        /// <param name="filePath">The absolute path to the project item
        /// being added.</param>
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
        }

        //---------------------------------------------------------------------
        // Other Methods
        //---------------------------------------------------------------------

        /// <summary>
        /// Cleans up Visual Studio solution and project directories.
        /// </summary>
        /// <param name="solutionRoot">Directory information for the Visual
        /// Studio solution root.</param>
        private void CleanSolution(DirectoryInfo solutionRoot)
        {
            DirectoryInfo solution = new DirectoryInfo(
                Path.Combine(solutionRoot.FullName, solutionRoot.Name));

            try
            {
                solution.Delete(true);
            }
            catch { }
        }
    }
}