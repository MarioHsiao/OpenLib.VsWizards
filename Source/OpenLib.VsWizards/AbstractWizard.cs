using Microsoft.VisualStudio.TemplateWizard;
using OpenLib.Utils;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;

namespace OpenLib.VsWizards
{
    /// <summary>
    /// The <c>AbstractWizard</c> type provides an abstract base class
    /// implementation containing common, reusable methods for implementing
    /// Visual Studio custom project template Wizard classes.
    /// </summary>
    public abstract class AbstractWizard
    {
        /// <summary>
        /// Defines the template variable name for the solution directory.
        /// </summary>
        protected const string VarSolutionDir = "$solutiondirectory$";
        
        /// <summary>
        /// Defines the template variable name for the wizard data.
        /// </summary>
        protected const string VarWizardData = "$wizarddata$";

        /// <summary>
        /// Defines the template variable name for the destination directory.
        /// </summary>
        protected const string VarDestinationDir = "$destinationdirectory$";

        /// <summary>
        /// Defines the template variable name for the safe project name.
        /// </summary>
        protected const string VarSafeProjectName = "$safeprojectname$";

        /// <summary>
        /// Defines the file extension for a Visual Studio template package.
        /// </summary>
        protected const string TemplatePackageExt = ".zip";

        /// <summary>
        /// Defines the namespace used for wizard template data XML.
        /// </summary>
        protected readonly XNamespace Ns = "http://schemas.microsoft.com/developer/vstemplate/2005";

        /// <summary>
        /// Gets a reference to the Visual Studio utilities.
        /// </summary>
        public IVsUtils VsUtils { get; internal set; }

        /// <summary>
        /// Gets directory information for the Visual Studio solution root.
        /// </summary>
        public DirectoryInfo SolutionRoot { get; internal set; }

        /// <summary>
        /// Gets a dictionary containing data for the template wizard.
        /// </summary>
        public Dictionary<string, string> WizardData { get; internal set; }

        /// <summary>
        /// Gets or sets the path to the Visual Studio project template.
        /// </summary>
        public string TemplatePath { get; internal set; }

        /// <summary>
        /// Gets the default destination path of the Visual Studio project to
        /// be created.
        /// </summary>
        public string DefaultDestinationPath { get; internal set; }

        /// <summary>
        /// Gets the namespace for the Visual Studio project to be created.
        /// </summary>
        public string NameSpace { get; internal set; }

        /// <summary>
        /// Gets a value indicating if the project template is valid.
        /// </summary>
        public bool IsValid { get; internal set; }

        /// <summary>
        /// Creates a new instance of the <c>AbstractWizard</c> class.
        /// </summary>
        protected AbstractWizard()
        {
            this.VsUtils = new VsUtils(new Vs2012CsInfo());
        }

        /// <summary>
        /// Defines an abstract method for validating creation of a Visual
        /// Studio project using a project template from a Wizard class
        /// implementation.
        /// </summary>
        /// <returns>A value indicating if the project template is valid.</returns>
        protected abstract bool Validate();

        /// <summary>
        /// Defines an abstract base method that sets common properties when a
        /// Wizard implementation's RunStarted method is executed.
        /// </summary>
        /// <param name="automationObject">A reference to the automation object used by the Visual Studio project template wizard.</param>
        /// <param name="replacementsDictionary">A reference to a dictionary of template variables that can be modified to provide customization.</param>
        /// <param name="runKind">Defines the type of the template the template wizard is creating.</param>
        /// <param name="customParams">An array of custom parameters passed to the template wizard.</param>
        protected void RunStarted
            (
                object automationObject,
                Dictionary<string, string> replacementsDictionary,
                WizardRunKind runKind,
                object[] customParams
            )
        {
            this.SolutionRoot = new DirectoryInfo(replacementsDictionary[VarSolutionDir]);

            this.WizardData = this.GetWizardData(replacementsDictionary);

            this.TemplatePath = customParams[0].ToString();
            this.DefaultDestinationPath = replacementsDictionary[VarDestinationDir];
            this.NameSpace = replacementsDictionary[VarSafeProjectName];

            this.IsValid = this.Validate();
        }

        /// <summary>
        /// Gets a dictionary containing data for the template wizard using
        /// data parsed from the specified replacements dictionary.
        /// </summary>
        /// <param name="replacementsDictionary">A reference to a dictionary of template variables that can be modified to provide customization.</param>
        /// <returns>A dictionary containing data for the template wizard.</returns>
        protected Dictionary<string, string> GetWizardData
            (
                Dictionary<string, string> replacementsDictionary
            )
        {
            Dictionary<string, string> wizardData = null;

            if (replacementsDictionary.ContainsKey(VarWizardData))
            {
                string data = replacementsDictionary[VarWizardData];

                if (!string.IsNullOrWhiteSpace(data))
                {
                    XElement root = XElement.Parse(data);

                    if (root != null && root.HasElements)
                    {
                        wizardData = (from e in root.Descendants(Ns + "entry")
                                      where e != null && e.HasAttributes && !e.IsEmpty
                                      select e
                                     ).ToDictionary(e => e.Attribute("name").Value,
                                                    e => e.HasElements ?
                                                        e.FirstNode.ToString() : e.Value);
                    }
                }
            }

            return wizardData;
        }
    }
} 
