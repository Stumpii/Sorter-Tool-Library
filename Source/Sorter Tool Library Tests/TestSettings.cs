using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newtonsoft.Json;
using SorterToolLibrary;
using System;
using System.IO;

namespace Settings
{
    [TestClass()]
    public class TestSettings
    {
        private const string ProductName = "CBOS Converter";

        #region Methods

        [TestMethod()]
        public void GetProgramSettingsTest()
        {
            HelperSettings helperSettings = GetHelperSettings();

            bool res = File.Exists(helperSettings.DRHMIConfigFilePath);
            Assert.IsTrue(res, $"DRHMIConfigFilePath file does not exist: {helperSettings.DRHMIConfigFilePath}.");

            FileInfo fileinfo = new FileInfo(helperSettings.WinCC74OutputTagFilePath);
            res = Directory.Exists(fileinfo.DirectoryName);
            Assert.IsTrue(res, $"WinCC74OutputTagFilePath folder does not exist: {helperSettings.WinCC74OutputTagFilePath}.");

            res = Directory.Exists(helperSettings.OutputTemplateFolder);
            Assert.IsTrue(res, $"OutputTemplateFolder folder does not exist: {helperSettings.OutputTemplateFolder}.");
        }

        public HelperSettings GetHelperSettings()
        {
            // Return hard coded settings
            return LatestTemplate(); // TEMPLATE TEST

            // deserialize JSON directly from a file

            /*
            string settingsFilename = "Program Settings - Fulcrum Sierra.json";
            string settingsFilename = "Program Settings - Petrobrazi.json";
            string settingsFilename = "Program Settings - Karish.json";
            */
            string settingsFilename = "Program Settings - Template.json";

            return GetHelperSettings(settingsFilename);
        }

        private HelperSettings LatestTemplate()
        {
            string projectFolder = @"C:\Programming\dotNet\D-R HMI Converter\Source\D-R HMI Converter\Templates\Input";

            return new HelperSettings
            {
                DRHMIConfigFilePath = Path.Combine(projectFolder, "D-R HMI Config (clean) Rev1.6.8.xls"),
                WonderwareOutputTagFilePath = Path.Combine(projectFolder, "D-R HMI Config (clean) Rev1.6.8.csv"),
                OutputTemplateFolder = @"C:\Programming\dotNet\D-R HMI Converter\Source\D-R HMI Converter\Templates\Output"
            };
        }
        public static HelperSettings GetHelperSettings(string SettingsFile)
        {
            HelperSettings helperSettings = null;

            // Get the Resources folder under the test project
            string AppDataFolder = Path.GetFullPath(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"..\..\Resources"));

            if (!Directory.Exists(AppDataFolder))
                Directory.CreateDirectory(AppDataFolder);

            // deserialize JSON directly from a file
            string settingsFilename = Path.Combine(AppDataFolder, SettingsFile);

            if (File.Exists(settingsFilename))
            {
                using (StreamReader file = File.OpenText(settingsFilename))
                {
                    JsonSerializer serializer = new JsonSerializer();
                    //serializer.Error += Serializer_Error;
                    ProgramSettings programSettings = (ProgramSettings)serializer.Deserialize(file, typeof(ProgramSettings));
                    helperSettings = programSettings.HelperSettings;
                }
            }
            return helperSettings;
        }

        #endregion Methods
    }

    public class ProgramSettings
    {
        #region "Properties"

        // TODO - Remove this class and let DR Engine serialize and de-serialize itself with a given path
        public HelperSettings HelperSettings { get; set; } = new HelperSettings();

        #endregion "Properties"
    }
}