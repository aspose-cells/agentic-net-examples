using System;
using System.Globalization;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Configure LoadOptions to use a specific culture (German in this case)
        LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);
        loadOptions.CultureInfo = new CultureInfo("de-DE"); // German uses comma as decimal separator

        // Load the workbook from an existing XLSX file with the specified culture
        Workbook workbook = new Workbook("input.xlsx", loadOptions);

        // Apply custom globalization settings to the loaded workbook
        workbook.Settings.GlobalizationSettings = new CustomGlobalizationSettings();

        // Demonstrate the effect of the custom globalization settings
        Cells cells = workbook.Worksheets[0].Cells;

        // Boolean values will be displayed using the overridden strings
        cells[0, 0].PutValue(true);
        cells[0, 1].PutValue(false);

        // An error value to show error string localization (optional)
        cells[0, 2].PutValue("#DIV/0!");

        Console.WriteLine($"Cell[0,0] (True):  {cells[0, 0].StringValue}");
        Console.WriteLine($"Cell[0,1] (False): {cells[0, 1].StringValue}");
        Console.WriteLine($"Cell[0,2] (Error): {cells[0, 2].StringValue}");

        // Save the workbook with the applied localization
        workbook.Save("output.xlsx");
    }

    // Custom globalization settings derived from GlobalizationSettings
    class CustomGlobalizationSettings : GlobalizationSettings
    {
        // Override boolean display strings (German)
        public override string GetBooleanValueString(bool bv)
        {
            return bv ? "WAHR" : "FALSCH";
        }

        // Optionally override error string translation
        public override string GetErrorValueString(string err)
        {
            // Example: translate DIV/0! error to German equivalent
            if (err == "#DIV/0!") return "#DIV/0!"; // could be "#DIV/0!" or another translation
            return base.GetErrorValueString(err);
        }
    }
}