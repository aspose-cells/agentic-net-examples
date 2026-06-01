using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Custom globalization settings that provide localized error messages
    public class CustomErrorGlobalizationSettings : GlobalizationSettings
    {
        // Override to map default error strings to custom localized strings
        public override string GetErrorValueString(string err)
        {
            // Map specific Excel error codes to custom messages
            return err switch
            {
                "#DIV/0!" => "Custom Division Error",
                "#VALUE!" => "Custom Type Mismatch",
                "#NAME?" => "Custom Identifier Error",
                "#N/A" => "Custom Not Available",
                "#REF!" => "Custom Reference Error",
                _ => base.GetErrorValueString(err) // Fallback to default behavior
            };
        }
    }

    public class GlobalizationSettingsGetErrorValueStringDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Apply the custom globalization settings to the workbook
                workbook.Settings.GlobalizationSettings = new CustomErrorGlobalizationSettings();

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];
                Cells cells = worksheet.Cells;

                // Create several error scenarios
                cells["A1"].Formula = "=1/0";                     // #DIV/0!
                cells["A2"].Formula = "=SUM(\"text\")";          // #VALUE!
                cells["A3"].Formula = "=UNKNOWNFUNC()";          // #NAME?
                cells["A4"].Formula = "=VLOOKUP(1,B1:C1,2,FALSE)"; // #N/A (if not found)
                cells["A5"].Formula = "=INDIRECT(\"Z1000\")";    // #REF!

                // Calculate formulas to generate the errors
                workbook.CalculateFormula();

                // Display the custom error strings for each cell
                Console.WriteLine($"A1 error display: {cells["A1"].DisplayStringValue}");
                Console.WriteLine($"A2 error display: {cells["A2"].DisplayStringValue}");
                Console.WriteLine($"A3 error display: {cells["A3"].DisplayStringValue}");
                Console.WriteLine($"A4 error display: {cells["A4"].DisplayStringValue}");
                Console.WriteLine($"A5 error display: {cells["A5"].DisplayStringValue}");

                // Save the workbook to verify that custom settings do not affect file content
                workbook.Save("CustomErrorGlobalizationDemo.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public static class Program
    {
        public static void Main()
        {
            GlobalizationSettingsGetErrorValueStringDemo.Run();
        }
    }
}