// Title: Runtime Localization Switch in Aspose.Cells for .NET (no workbook reload)
// Description: Demonstrates how to change a workbook's language, region, and culture on the fly using Aspose.Cells. The example sets an initial English locale, calculates a formula, then calls a SwitchLocalization method that updates Workbook.Settings (LanguageCode, Region, CultureInfo) and applies SettableGlobalizationSettings to adjust list separators and map function names (e.g., SUM → SUMME). Formulas can be entered in the new language and recalculated without reopening the file.
// Keywords: Aspose.Cells runtime localization | C# change workbook language | SettableGlobalizationSettings example | localized Excel functions .NET | culture-specific list separator | dynamic region switch Aspose.Cells | German Excel functions SUMME | French Excel functions SOMME | multi‑regional reporting C# | Aspose.Cells workbook Settings.LanguageCode
// Common Searches: change workbook locale at runtime Aspose.Cells | map Excel function names to local language C# | adjust list separator for German culture Aspose.Cells | switch Excel region without reloading file | SettableGlobalizationSettings usage example
// Developer Intent: Modify a workbook’s locale and function mappings while it remains open.
// Use Cases: Allow end‑users to toggle between English (US) and German interfaces, instantly reflecting localized formulas and number formats. | Generate financial statements for different countries in a single session, applying appropriate list separators and numeric conventions. | Extend the SwitchLocalization routine to support additional cultures such as French, Japanese, or Russian with custom function name maps.
// AI Prompts: Add Spanish locale support with localized function names to the SwitchLocalization method. | Create code that saves a user‑selected CountryCode and reapplies the same globalization settings when reopening a workbook. | Explain the impact of SettableGlobalizationSettings on formula parsing and calculation in Aspose.Cells.

using System;
using System.Globalization;
using Aspose.Cells;

namespace AsposeCellsLocalizationDemo
{
    // Demonstrates how to change a workbook's language, region, and culture on the fly using Aspose.Cells. The example sets an initial English locale, calculates a formula, then calls a SwitchLocalization method that updates Workbook.Settings (LanguageCode, Region, CultureInfo) and applies SettableGlobalizationSettings to adjust list separators and map function names (e.g., SUM → SUMME). Formulas can be entered in the new language and recalculated without reopening the file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and add sample data
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            cells["A1"].PutValue("Value");
            cells["A2"].PutValue(10);
            cells["A3"].PutValue(20);
            cells["A4"].PutValue(30);

            // Initial localization: English (USA)
            SwitchLocalization(workbook, CountryCode.USA);
            // Use standard function name
            cells["B1"].Formula = "=SUM(A2:A4)";
            workbook.CalculateFormula();
            Console.WriteLine("English locale - SUM result: " + cells["B1"].Value);
            Console.WriteLine("Number format (English): " + cells["B1"].StringValue);

            // Switch to German locale at runtime
            SwitchLocalization(workbook, CountryCode.Germany);
            // Use localized function name (SUMME) after mapping
            cells["C1"].Formula = "=SUMME(A2:A4)";
            workbook.CalculateFormula();
            Console.WriteLine("German locale - SUMME result: " + cells["C1"].Value);
            Console.WriteLine("Number format (German): " + cells["C1"].StringValue);

            // Save the workbook (no reload required)
            workbook.Save("LocalizationSwitchDemo.xlsx");
        }

        /// <param name="workbook">Target workbook.</param>
        /// <param name="country">Desired country/locale.</param>
        static void SwitchLocalization(Workbook workbook, CountryCode country)
        {
            // Set UI language and regional settings
            workbook.Settings.LanguageCode = country;
            workbook.Settings.Region = country;

            // Adjust CultureInfo based on the country code.
            // Most CountryCode enum names match .NET culture names (e.g., Germany -> "de-DE").
            // For simplicity, we map a few common ones; otherwise fallback to invariant culture.
            CultureInfo culture = country switch
            {
                CountryCode.USA => new CultureInfo("en-US"),
                CountryCode.Germany => new CultureInfo("de-DE"),
                CountryCode.France => new CultureInfo("fr-FR"),
                CountryCode.Japan => new CultureInfo("ja-JP"),
                CountryCode.China => new CultureInfo("zh-CN"),
                CountryCode.Russia => new CultureInfo("ru-RU"),
                _ => CultureInfo.InvariantCulture
            };
            workbook.Settings.CultureInfo = culture;

            // Create and apply custom globalization settings.
            // Example: map the SUM function to its local name.
            SettableGlobalizationSettings gSettings = new SettableGlobalizationSettings();

            // Map list separator according to culture (comma for English, semicolon for German, etc.)
            char listSep = culture.TextInfo.ListSeparator[0];
            gSettings.SetListSeparator(listSep);

            // Map standard function names to localized equivalents.
            // For German we map SUM -> SUMME, AVERAGE -> MITTELWERT, etc.
            if (country == CountryCode.Germany)
            {
                gSettings.SetLocalFunctionName("SUM", "SUMME", true);
                gSettings.SetLocalFunctionName("AVERAGE", "MITTELWERT", true);
            }
            else if (country == CountryCode.France)
            {
                gSettings.SetLocalFunctionName("SUM", "SOMME", true);
                gSettings.SetLocalFunctionName("AVERAGE", "MOYENNE", true);
            }
            else
            {
                // Default to English (no mapping needed)
                gSettings.SetLocalFunctionName("SUM", "SUM", true);
                gSettings.SetLocalFunctionName("AVERAGE", "AVERAGE", true);
            }

            // Apply the globalization settings to the workbook.
            workbook.Settings.GlobalizationSettings = gSettings;
        }
    }
}
