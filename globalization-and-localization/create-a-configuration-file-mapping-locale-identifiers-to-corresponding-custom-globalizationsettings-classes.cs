using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using Aspose.Cells;

namespace AsposeCellsLocaleConfig
{
    // Base class for custom globalization settings (inherits from Aspose.Cells.GlobalizationSettings)
    public abstract class CustomGlobalizationSettingsBase : GlobalizationSettings
    {
        // Common helper can be added here if needed
    }

    // Example custom settings for English (US)
    public class EnglishGlobalizationSettings : CustomGlobalizationSettingsBase
    {
        public override string GetBooleanValueString(bool bv)
        {
            return bv ? "TRUE_EN" : "FALSE_EN";
        }
    }

    // Example custom settings for Russian
    public class RussianGlobalizationSettings : CustomGlobalizationSettingsBase
    {
        public override string GetBooleanValueString(bool bv)
        {
            return bv ? "ИСТИНА" : "ЛОЖЬ";
        }

        public override string GetErrorValueString(string err)
        {
            return err switch
            {
                "#NAME?" => "#ИМЯ?",
                "#DIV/0!" => "#ДЕЛ/0!",
                "#REF!" => "#ССЫЛКА!",
                "#VALUE!" => "#ЗНАЧ?",
                "#N/A" => "#Н/Д",
                "#NUM!" => "#ЧИСЛО!",
                "#NULL!" => "#ПУСТО!",
                _ => base.GetErrorValueString(err)
            };
        }
    }

    // Factory that reads a configuration file and provides the appropriate GlobalizationSettings instance
    public static class GlobalizationSettingsFactory
    {
        // Mapping from LCID to the Type of the custom settings class
        private static readonly Dictionary<int, Type> _localeTypeMap = new Dictionary<int, Type>();

        // Load mapping from a JSON configuration file.
        // The JSON should be an object where keys are LCIDs (as strings) and values are the fully qualified type names.
        // Example: { "1033": "AsposeCellsLocaleConfig.EnglishGlobalizationSettings", "1049": "AsposeCellsLocaleConfig.RussianGlobalizationSettings" }
        public static void LoadConfiguration(string configFilePath)
        {
            if (!File.Exists(configFilePath))
                throw new FileNotFoundException($"Configuration file not found: {configFilePath}");

            string json = File.ReadAllText(configFilePath);
            var rawMap = JsonSerializer.Deserialize<Dictionary<string, string>>(json);

            _localeTypeMap.Clear();

            foreach (var kvp in rawMap)
            {
                if (!int.TryParse(kvp.Key, out int lcid))
                    continue; // skip invalid keys

                // Resolve the type by name
                Type type = Type.GetType(kvp.Value, throwOnError: false, ignoreCase: true);
                if (type != null && typeof(GlobalizationSettings).IsAssignableFrom(type))
                {
                    _localeTypeMap[lcid] = type;
                }
            }
        }

        // Retrieve an instance of the appropriate GlobalizationSettings for the given LCID.
        // If no custom mapping exists, returns the default GlobalizationSettings instance.
        public static GlobalizationSettings GetSettings(int lcid)
        {
            if (_localeTypeMap.TryGetValue(lcid, out Type settingsType))
            {
                // Use Activator to create an instance (must have a parameterless constructor)
                return (GlobalizationSettings)Activator.CreateInstance(settingsType);
            }

            // Fallback to default settings
            return new GlobalizationSettings();
        }
    }

    // Demonstration of using the factory to apply locale‑specific globalization settings to a workbook
    public class Demo
    {
        public static void Run()
        {
            try
            {
                // Path to the JSON configuration file
                string configPath = "localeConfig.json";

                // If the config file does not exist, create a default one
                if (!File.Exists(configPath))
                {
                    var defaultConfig = new Dictionary<string, string>
                    {
                        { "1033", "AsposeCellsLocaleConfig.EnglishGlobalizationSettings" },
                        { "1049", "AsposeCellsLocaleConfig.RussianGlobalizationSettings" }
                    };
                    string defaultJson = JsonSerializer.Serialize(defaultConfig, new JsonSerializerOptions { WriteIndented = true });
                    File.WriteAllText(configPath, defaultJson);
                }

                // Load the mapping (ensure the file exists with proper content)
                GlobalizationSettingsFactory.LoadConfiguration(configPath);

                // Example: choose a locale identifier (LCID). 1049 = Russian, 1033 = English (US)
                int localeId = 1049;

                // Obtain the appropriate globalization settings instance
                GlobalizationSettings localeSettings = GlobalizationSettingsFactory.GetSettings(localeId);

                // Create a new workbook and apply the settings
                Workbook wb = new Workbook();
                wb.Settings.GlobalizationSettings = localeSettings;

                // Populate some sample data to illustrate the effect
                Worksheet ws = wb.Worksheets[0];
                ws.Cells["A1"].PutValue(true);               // Boolean value
                ws.Cells["A2"].PutValue("#DIV/0!");          // Error value

                // Save the workbook
                string outputPath = "LocalizedWorkbook.xlsx";
                wb.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during demo execution: {ex.Message}");
            }
        }
    }

    // Entry point
    class Program
    {
        static void Main()
        {
            Demo.Run();
        }
    }
}