// Title: Custom GlobalizationSettings in Aspose.Cells .NET – Map Excel Functions to Localized Names
// Description: Demonstrates how to subclass Aspose.Cells.GlobalizationSettings, override GetLocalFunctionName to translate standard formulas (e.g., SUM, AVERAGE) into custom names (LOCALSUM, LOCALAVERAGE), apply the settings to a Workbook, use the localized formulas, calculate results, and save the file.
// Keywords: Aspose.Cells | Custom GlobalizationSettings | GetLocalFunctionName override | localized Excel functions | C# .NET | Excel formula localization | map SUM to LOCALSUM | custom function names in formulas | workbook globalization | culture invariant workbook
// Common Searches: override GetLocalFunctionName Aspose.Cells | custom function names for Excel formulas .NET | apply GlobalizationSettings to workbook | map Excel functions to localized names C# | Aspose.Cells localized formulas example
// Developer Intent: Create a subclass of GlobalizationSettings that provides custom localized function names and integrate it into a workbook to generate spreadsheets with language‑specific formulas.
// Use Cases: Generate Excel files for non‑English locales where function names differ. | Maintain a single code base while supporting multiple regional formula conventions. | Ensure formulas are correctly interpreted by end‑users who expect local function names.
// AI Prompts: Write a CustomGlobalizationSettings class that also maps MIN, MAX, and COUNT to LOCALMIN, LOCALMAX, and LOCALCOUNT. | Show how to retrieve the original English function name from a localized name using Aspose.Cells. | Provide code to toggle between default and custom globalization settings at runtime based on user locale.

using System;
using System.Globalization;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Custom globalization settings that maps standard function names to localized ones
    // Demonstrates how to subclass Aspose.Cells.GlobalizationSettings, override GetLocalFunctionName to translate standard formulas (e.g., SUM, AVERAGE) into custom names (LOCALSUM, LOCALAVERAGE), apply the settings to a Workbook, use the localized formulas, calculate results, and save the file.
    public class CustomGlobalizationSettings : GlobalizationSettings
    {
        // Override to provide custom local function names
        public override string GetLocalFunctionName(string standardName)
        {
            // Example: map SUM to LOCALSUM and AVERAGE to LOCALAVERAGE
            if (standardName.Equals("SUM", StringComparison.OrdinalIgnoreCase))
                return "LOCALSUM";
            if (standardName.Equals("AVERAGE", StringComparison.OrdinalIgnoreCase))
                return "LOCALAVERAGE";

            // Fallback to base implementation for other functions
            return base.GetLocalFunctionName(standardName);
        }
    }

    public class GlobalizationSettingsDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook (lifecycle create)
                Workbook workbook = new Workbook();

                // Ensure the workbook uses an invariant culture before applying custom globalization
                workbook.Settings.CultureInfo = CultureInfo.InvariantCulture;

                // Assign the custom globalization settings to the workbook
                workbook.Settings.GlobalizationSettings = new CustomGlobalizationSettings();

                Worksheet worksheet = workbook.Worksheets[0];

                // Populate sample data
                worksheet.Cells["B1"].PutValue(5);
                worksheet.Cells["B2"].PutValue(15);
                worksheet.Cells["C1"].PutValue(2);
                worksheet.Cells["C2"].PutValue(4);

                // Use the localized function name in a formula
                Cell sumCell = worksheet.Cells["B3"];
                sumCell.Formula = "=LOCALSUM(B1:B2)";

                // Use another localized function name
                Cell avgCell = worksheet.Cells["C3"];
                avgCell.Formula = "=LOCALAVERAGE(C1:C2)";

                // Calculate all formulas (lifecycle load/compute)
                workbook.CalculateFormula();

                // Output the results
                Console.WriteLine($"Result of LOCALSUM: {sumCell.DoubleValue}");
                Console.WriteLine($"Result of LOCALAVERAGE: {avgCell.DoubleValue}");

                // Save the workbook (lifecycle save)
                string outputPath = "CustomGlobalizationDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point
    class Program
    {
        static void Main()
        {
            GlobalizationSettingsDemo.Run();
        }
    }
}
