using System;
using Aspose.Cells;
using System.Globalization;

class LocalizationSwitcher
{
    static void Main()
    {
        // Create a new workbook and add sample data
        Workbook wb = new Workbook();
        Worksheet ws = wb.Worksheets[0];
        ws.Cells["A1"].PutValue(10);
        ws.Cells["A2"].PutValue(20);
        ws.Cells["A3"].PutValue(30);

        // Use the default (English) function name
        ws.Cells["B1"].Formula = "=SUM(A1:A3)";
        wb.CalculateFormula();
        Console.WriteLine("Result with English SUM: " + ws.Cells["B1"].Value);

        // Switch localization to German at runtime (no reload)
        SwitchLocalization(wb, CountryCode.Germany);

        // Use the German function name after switching
        ws.Cells["B2"].Formula = "=SUMME(A1:A3)"; // German name for SUM
        wb.CalculateFormula();
        Console.WriteLine("Result with German SUMME: " + ws.Cells["B2"].Value);

        // Save the workbook (uses the provided save rule)
        wb.Save("LocalizationSwitchDemo.xlsx");
    }

    // Changes UI language, region, and function name mappings for the given workbook
    static void SwitchLocalization(Workbook workbook, CountryCode country)
    {
        // Set UI language and regional settings
        workbook.Settings.LanguageCode = country;
        workbook.Settings.Region = country;

        // Create globalization settings that allow custom function name mapping
        SettableGlobalizationSettings gSettings = new SettableGlobalizationSettings();

        // Define mappings for supported languages
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
        else if (country == CountryCode.Italy)
        {
            gSettings.SetLocalFunctionName("SUM", "SOMMA", true);
            gSettings.SetLocalFunctionName("AVERAGE", "MEDIA", true);
        }

        // Apply the custom globalization settings to the workbook
        workbook.Settings.GlobalizationSettings = gSettings;
    }
}