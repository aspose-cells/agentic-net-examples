// Title: Validate French localized boolean strings from Excel formula cells using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells retrieve French display string for TRUE() formula | Validate that =TRUE() shows VRAI in a French workbook with Aspose.Cells | How to check localized FALSE() value after formula calculation in Aspose.Cells .NET | Aspose.Cells DisplayString for boolean formulas when CultureInfo is fr-FR | Testing French boolean localization in Excel cells using Aspose.Cells API
// Tags: Aspose.Cells Boolean localization | C# CellValueFormatStrategy DisplayString | Workbook Settings CultureInfo French | Excel formula TRUE FALSE localization | Validate localized boolean strings Aspose.Cells

using System;
using System.Globalization;
using Aspose.Cells;

// The example creates a Workbook, sets its Settings.CultureInfo to French (fr-FR), writes =TRUE() and =FALSE() formulas into cells A1 and A2, calculates the workbook, retrieves the displayed strings with CellValueFormatStrategy.DisplayString, and verifies they match the French terms "VRAI" and "FAUX", throwing an exception on mismatch and confirming success otherwise.
class BooleanLocalizationValidator
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Set the workbook culture to French (France) to get localized boolean strings
            workbook.Settings.CultureInfo = new CultureInfo("fr-FR");

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Cell A1: formula that evaluates to TRUE
            Cell trueCell = sheet.Cells["A1"];
            trueCell.Formula = "=TRUE()";

            // Cell A2: formula that evaluates to FALSE
            Cell falseCell = sheet.Cells["A2"];
            falseCell.Formula = "=FALSE()";

            // Calculate all formulas in the workbook
            workbook.CalculateFormula();

            // Retrieve the displayed (localized) string values using the DisplayString format strategy
            string trueDisplay = trueCell.GetStringValue(CellValueFormatStrategy.DisplayString);
            string falseDisplay = falseCell.GetStringValue(CellValueFormatStrategy.DisplayString);

            // Validate the localized strings
            if (trueDisplay != "VRAI")
                throw new Exception($"Localized TRUE value mismatch. Expected 'VRAI', got '{trueDisplay}'.");

            if (falseDisplay != "FAUX")
                throw new Exception($"Localized FALSE value mismatch. Expected 'FAUX', got '{falseDisplay}'.");

            // If no exception was thrown, validation succeeded
            Console.WriteLine("Boolean localization validation passed.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
