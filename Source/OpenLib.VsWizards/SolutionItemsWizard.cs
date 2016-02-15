using EnvDTE;
using Microsoft.VisualStudio.TemplateWizard;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;

namespace OpenLib.VsWizards
{
    /// <summary>
    /// The <c>SolutionItemsWizard</c> type provides an implementation of the
    /// <c>IWizard</c> interface that provides a custom Visual Studio project
    /// template wizard for adding solution items to a Visual Studio solution.
    /// </summary>
    /// <remarks>
    /// This class inherits from AbstractWizard, which provides common
    /// Wizard properties and methods.
    /// </remarks>
    public class SolutionItemsWizard : AbstractWizard, IWizard
    {
        //---------------------------------------------------------------------
        // Fields
        //---------------------------------------------------------------------

        /// <summary>
        /// Defines the wizard data dictionary key for the solution items.
        /// </summary>
        private const string WizardDataKeySolutionItems = "solutionItems";

        /// <summary>
        /// Defines the physical Visual Studio solution root.
        /// </summary>
        private const string PhysicalSolutionRoot = "Root";

        //---------------------------------------------------------------------
        // Constructors
        //---------------------------------------------------------------------

        /// <summary>
        /// Creates a new instance of the <c>SolutionItemsWizard</c> class.
        /// </summary>
        public SolutionItemsWizard() : base()
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
        /// <param name="projectItem">A reference to the project item being added
        /// to the Visual Studio project.</param>
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
            this.AddSolutionItems(
                this.SolutionRoot,
                this.TemplatePath,
                this.GetSolutionItems(this.WizardData));
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

            this.TemplatePath = Path.GetDirectoryName(
                Path.GetDirectoryName(this.TemplatePath));
        }

        //---------------------------------------------------------------------
        // Other Methods
        //---------------------------------------------------------------------

        /// <summary>
        /// Gets a list of all solution items to be added to a Visual Studio
        /// solution.
        /// </summary>
        /// <param name="wizardData">A dictionary containing data for the
        /// template wizard.</param>
        /// <returns>A list of all solution items to be added to a Visual
        /// Studio solution.</returns>
        /// <remarks>
        /// Each list item is stored as a Tuple:
        /// Item1 = Physical folder
        /// Item2 = Solution folder
        /// Item3 = Solution item file
        /// </remarks>
        private List<Tuple<string, string, string>> GetSolutionItems(Dictionary<string, string> wizardData)
        {
            List<Tuple<string, string, string>> solutionItems = null;
            string data = wizardData[WizardDataKeySolutionItems];

            if (!string.IsNullOrWhiteSpace(data))
            {
                XElement root = XElement.Parse(data);

                if (root != null && root.HasElements)
                {
                    solutionItems = (from e in root.Descendants(Ns + "solutionItem")
                                     where e != null && e.HasElements &&
                                        e.Element(Ns + "physicalFolder") != null &&
                                        e.Element(Ns + "solutionFolder") != null &&
                                        e.Element(Ns + "solutionItemFile") != null
                                     select new Tuple<string, string, string>
                                        (
                                            e.Element(Ns + "physicalFolder").Value,
                                            e.Element(Ns + "solutionFolder").Value,
                                            e.Element(Ns + "solutionItemFile").Value
                                        )
                                    ).ToList();
                }
            }

            return solutionItems;
        }

        /// <summary>
        /// Adds Visual Studio solution items to solution folders using the
        /// specified solution item metadata and creates the files in their
        /// appropriate physical folders on the file system.
        /// </summary>
        /// <param name="solutionRoot">Directory information for the Visual
        /// Studio solution root.</param>
        /// <param name="templatePath">The path to the Visual Studio project
        /// template.</param>
        /// <param name="solutionItems">A list of metadata for creation of
        /// the solution items.</param>
        private void AddSolutionItems(DirectoryInfo solutionRoot,
                                      string templatePath,
                                      List<Tuple<string, string, string>> solutionItems)
        {
            if (solutionItems != null && solutionItems.Count > 0)
            {
                solutionItems.ForEach(si =>
                    {
                        Project solutionFolder = this.VsUtils.FindSolutionFolder(si.Item2);

                        if (solutionFolder != null)
                        {
                            string safePath = this.GetSafePath(si.Item2);
                            string safeTemplatePath = string.Concat(safePath, TemplatePackageExt);

                            FileInfo sourcePath = new FileInfo(
                                Path.Combine(templatePath, safeTemplatePath, si.Item3));

                            FileInfo destinationPath = new FileInfo(
                                (!si.Item1.Equals(PhysicalSolutionRoot)) ?
                                Path.Combine(solutionRoot.FullName, si.Item1, si.Item3) :
                                Path.Combine(solutionRoot.FullName, si.Item3));

                            if (sourcePath.Exists)
                            {
                                DirectoryInfo destinationDir = new DirectoryInfo(
                                    Path.GetDirectoryName(destinationPath.FullName));

                                if (!destinationDir.Exists)
                                    destinationDir.Create();

                                sourcePath.CopyTo(destinationPath.FullName, true);
                                solutionFolder.ProjectItems.AddFromFile(destinationPath.FullName);
                            }
                        }
                    });
            }
        }

        /// <summary>
        /// Gets a safe path by removing invalid path characters.
        /// </summary>
        /// <param name="path">The path in which to cleanse of invalid
        /// characters.</param>
        /// <returns>A safe path by removing invalid path characters.</returns>
        private string GetSafePath(string path)
        {
            return path.Replace(" ", "").Replace(".", "");
        }
    }
}
