// Title: Create a custom GlobalizationSettings class in C# to localize Excel formula function names with Aspose.Cells
// Common Searches: C# Aspose.Cells override GetLocalFunctionName to translate Excel functions to Portuguese | How to map SUM to SOMA using custom GlobalizationSettings in Aspose.Cells | Assign custom GlobalizationSettings to a Workbook for localized formulas | Example of Excel formula localization with Aspose.Cells and custom globalization settings | Save workbook with localized function names using Aspose.Cells C# API
// Tags: override GetLocalFunctionName Aspose.Cells | custom GlobalizationSettings C# | localize Excel formula functions Aspose.Cells | Portuguese function name mapping Aspose.Cells | assign GlobalizationSettings to Workbook

using System;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // Custom globalization settings that overrides function names for target language
    // Shows how to implement a CustomGlobalizationSettings class that overrides GetLocalFunctionName to map English Excel functions (SUM, AVERAGE, MAX, MIN) to Portuguese equivalents, assign this settings object to a Workbook, add a formula using the English name, and save the file so the formula is stored with the localized function name.
    public class CustomGlobalizationSettings : GlobalizationSettings
    {
        // Override the method that returns the localized function name
        public override string GetLocalFunctionName(string functionName)
        {
            // Example: map English function names to Portuguese equivalents
            // Add more mappings as needed
            switch (functionName.ToUpperInvariant())
            {
                case "SUM":
                    return "SOMA";
                case "AVERAGE":
                    return "MÉDIA";
                case "MAX":
                    return "MÁXIMO";
                case "MIN":
                    return "MÍNIMO";
                default:
                    // Fallback to the default implementation for unmapped functions
                    return base.GetLocalFunctionName(functionName);
            }
        }
    }

    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Assign the custom globalization settings to the workbook
            workbook.Settings.GlobalizationSettings = new CustomGlobalizationSettings();

            // Example: add a formula using the English function name; it will be localized when saved
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue(10);
            sheet.Cells["A2"].PutValue(20);
            sheet.Cells["A3"].Formula = "SUM(A1:A2)";

            // Save the workbook (the formula will be stored using the localized name)
            workbook.Save("CustomGlobalization.xlsx");
        }
    }
}
