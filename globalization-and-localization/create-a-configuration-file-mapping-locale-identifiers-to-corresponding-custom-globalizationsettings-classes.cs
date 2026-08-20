// Title: C# Mapping LCID to Custom GlobalizationSettings for Aspose.Cells Workbooks
// Description: Demonstrates how to create locale‑specific GlobalizationSettings (e.g., Russian and German) in Aspose.Cells, store them in an LCID‑keyed dictionary, and retrieve the appropriate settings via a factory method that falls back to the default when a locale is not mapped.
// Keywords: Aspose.Cells | Custom GlobalizationSettings | LCID mapping | C# localization | Russian Excel workbook | German Excel workbook | locale factory | Excel globalization | culture specific error strings | boolean translation
// Common Searches: Aspose.Cells map LCID to GlobalizationSettings | C# custom globalization for Russian Excel | German error messages in Aspose.Cells | How to localize boolean values in Aspose.Cells | Factory pattern for locale settings Aspose.Cells | Add new locale to Aspose.Cells globalization
// Developer Intent: Configure Aspose.Cells to apply custom globalization (boolean and error text) based on a given locale identifier.
// Use Cases: Generate a workbook for Russian users where TRUE/FALSE and error codes appear in Russian by using LCID 1049. | Switch between German and default globalization at runtime by passing LCID 1031 to the GlobalizationSettingsFactory. | Extend the mapping dictionary to support additional locales (e.g., French, Spanish) without modifying workbook creation code.
// AI Prompts: Write C# code that adds a FrenchGlobalizationSettings class with appropriate translations and updates the GlobalizationSettingsFactory mapping. | Show how to obtain the current thread's CultureInfo LCID and apply the matching GlobalizationSettings to an Aspose.Cells workbook. | Create a unit test that verifies GlobalizationSettingsFactory returns RussianGlobalizationSettings for LCID 1049 and GermanGlobalizationSettings for LCID 1031.

using System;
using System.Collections.Generic;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsLocaleConfig
{
    // Custom globalization for Russian locale (LCID 1049)
    // Demonstrates how to create locale‑specific GlobalizationSettings (e.g., Russian and German) in Aspose.Cells, store them in an LCID‑keyed dictionary, and retrieve the appropriate settings via a factory method that falls back to the default when a locale is not mapped.
    public class RussianGlobalizationSettings : GlobalizationSettings
    {
        public override string GetBooleanValueString(bool value)
        {
            return value ? "ИСТИНА" : "ЛОЖЬ";
        }

        public override string GetErrorValueString(string err)
        {
            // Map standard error strings to Russian equivalents
            return err switch
            {
                "#NAME?" => "#ИМЯ?",
                "#DIV/0!" => "#ДЕЛ/0!",
                "#REF!" => "#ССЫЛКА!",
                "#VALUE!" => "#ЗНАЧ!",
                "#N/A" => "#Н/Д",
                "#NUM!" => "#ЧИСЛО!",
                "#NULL!" => "#ПУСТО!",
                _ => base.GetErrorValueString(err)
            };
        }
    }

    // Custom globalization for German locale (LCID 1031)
    public class GermanGlobalizationSettings : GlobalizationSettings
    {
        public override string GetBooleanValueString(bool value)
        {
            return value ? "WAHR" : "FALSCH";
        }

        public override string GetErrorValueString(string err)
        {
            // Example mapping for German; extend as needed
            return err switch
            {
                "#NAME?" => "#NAME?",
                "#DIV/0!" => "#DIV/0!",
                "#REF!" => "#BEZUG!",
                "#VALUE!" => "#WERT!",
                "#N/A" => "#NV",
                "#NUM!" => "#ZAHL!",
                "#NULL!" => "#NULL!",
                _ => base.GetErrorValueString(err)
            };
        }
    }

    // Factory that returns a GlobalizationSettings instance based on LCID
    public static class GlobalizationSettingsFactory
    {
        // Mapping of locale identifiers (LCID) to corresponding settings instances
        private static readonly Dictionary<int, GlobalizationSettings> _settingsMap = new()
        {
            // 1049 = Russian (Russia)
            { 1049, new RussianGlobalizationSettings() },

            // 1031 = German (Germany)
            { 1031, new GermanGlobalizationSettings() },

            // Add more mappings as required
        };

        // Returns the appropriate settings; if not found, returns default settings
        public static GlobalizationSettings GetSettings(int lcid)
        {
            if (_settingsMap.TryGetValue(lcid, out GlobalizationSettings settings))
            {
                return settings;
            }

            // Fallback to default (no customization)
            return new GlobalizationSettings();
        }
    }

    // Demonstration of applying the configuration to a workbook
    public class Program
    {
        public static void Main()
        {
            // Example locale identifier; change to test different locales
            int localeId = 1049; // Russian

            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate some sample data
            cells[0, 0].PutValue(true);   // Boolean true
            cells[0, 1].PutValue(false);  // Boolean false
            cells[0, 2].PutValue("#DIV/0!"); // Error string

            // Retrieve the custom globalization settings for the given locale
            GlobalizationSettings localeSettings = GlobalizationSettingsFactory.GetSettings(localeId);

            // Apply the settings to the workbook
            workbook.Settings.GlobalizationSettings = localeSettings;

            // Demonstrate that the settings affect cell string values
            for (int col = 0; col < 3; col++)
            {
                Console.WriteLine($"Cell[0,{col}]: {cells[0, col].StringValue}");
            }

            // Save the workbook (output file name reflects the locale)
            string fileName = $"Workbook_LCID_{localeId}.xlsx";
            workbook.Save(fileName);
            Console.WriteLine($"Workbook saved as {fileName}");
        }
    }
}
