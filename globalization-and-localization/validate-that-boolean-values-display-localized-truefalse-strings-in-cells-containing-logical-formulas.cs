// Title: Validate French “VRAI”/“FAUX” display for =TRUE() and =FALSE() formulas in an Aspose.Cells C# workbook
// AI Prompts: Create C# code using Aspose.Cells that sets Workbook.Settings.CultureInfo to "fr-FR", writes the formulas =TRUE() and =FALSE() into cells, runs workbook.CalculateFormula(), and captures each cell's StringValue. | Show how to compare the StringValue of a calculated Boolean cell with the expected French terms "VRAI" and "FAUX" and log whether the localization is correct.
// Common Searches: Aspose.Cells C# how to get French localized result for =TRUE() formula | verify boolean formula localization in an Excel workbook with Aspose.Cells | set workbook culture to fr-FR and read localized boolean strings using Aspose.Cells | C# Aspose.Cells calculate formulas and retrieve StringValue for VRAI/FAUX | display VRAI and FAUX instead of TRUE/FALSE in Aspose.Cells workbook
// Tags: set workbook cultureinfo Aspose.Cells | calculate formulas Aspose.Cells C# | read stringvalue localized boolean | validate french boolean localization | save workbook as xlsx Aspose.Cells

using System;
using System.Globalization;
using Aspose.Cells;

// The program creates a new Workbook, sets its culture to French (fr-FR), inserts =TRUE() and =FALSE() formulas into cells A1 and A2, calculates the formulas, reads the displayed StringValue for each cell, checks that they match the French words "VRAI" and "FAUX", outputs the validation results, and saves the workbook as an XLSX file.
class BooleanLocalizationValidator
{
    static void Main()
    {
        try
        {
            // Create a new workbook (rule: create)
            Workbook workbook = new Workbook();

            // Set the workbook culture to French (localized true/false)
            workbook.Settings.CultureInfo = new CultureInfo("fr-FR");

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Insert logical formulas
            sheet.Cells["A1"].Formula = "=TRUE()";
            sheet.Cells["A2"].Formula = "=FALSE()";

            // Calculate the formulas
            workbook.CalculateFormula();

            // Retrieve displayed string values (localized)
            string displayedTrue = sheet.Cells["A1"].StringValue;   // Expected: "VRAI"
            string displayedFalse = sheet.Cells["A2"].StringValue; // Expected: "FAUX"

            // Validate the localization
            bool isTrueLocalized = displayedTrue.Equals("VRAI", StringComparison.OrdinalIgnoreCase);
            bool isFalseLocalized = displayedFalse.Equals("FAUX", StringComparison.OrdinalIgnoreCase);

            Console.WriteLine($"Cell A1 displays '{displayedTrue}' - Localization valid: {isTrueLocalized}");
            Console.WriteLine($"Cell A2 displays '{displayedFalse}' - Localization valid: {isFalseLocalized}");

            // Save the workbook (rule: save)
            string outputPath = "BooleanLocalizationDemo.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
