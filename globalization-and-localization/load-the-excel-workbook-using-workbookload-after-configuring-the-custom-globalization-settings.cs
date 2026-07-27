using System;
using System.Globalization;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Configure load options with a specific culture (German in this example)
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.CultureInfo = new CultureInfo("de-DE");

        // Load the workbook using the configured LoadOptions
        Workbook workbook = new Workbook("input.xlsx", loadOptions);

        // Apply custom globalization settings to the loaded workbook
        workbook.Settings.GlobalizationSettings = new CustomGlobalizationSettings();

        // Example usage: display a cell value to see the effect of the culture settings
        Console.WriteLine("Cell A1 value: " + workbook.Worksheets[0].Cells["A1"].StringValue);

        // Save the workbook after applying the custom settings
        workbook.Save("output.xlsx");
    }

    // Custom globalization settings example
    private class CustomGlobalizationSettings : GlobalizationSettings
    {
        // Override boolean value strings
        public override string GetBooleanValueString(bool value)
        {
            return value ? "ИСТИНА" : "ЛОЖЬ";
        }

        // Override error value strings (example translation)
        public override string GetErrorValueString(string error)
        {
            return error switch
            {
                "#DIV/0!" => "#ДЕЛ/0!",
                "#NAME?" => "#ИМЯ?",
                _ => base.GetErrorValueString(error)
            };
        }
    }
}