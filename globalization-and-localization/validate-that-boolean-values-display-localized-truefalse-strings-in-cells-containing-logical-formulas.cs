// Title: Verify German localized TRUE and FALSE strings for logical formulas with Aspose.Cells in C#
// Common Searches: Aspose.Cells C# display German WAHR and FALSCH from =TRUE() and =FALSE() formulas | set workbook CultureInfo to de-DE and read StringValue of boolean formulas in C# | validate that calculated cells return localized boolean text using Aspose.Cells | how to get localized boolean string values after workbook.CalculateFormula in C# | check Excel boolean localization with Aspose.Cells when saving as .xlsx
// Tags: Aspose.Cells workbook culture localization | C# calculate formulas with CultureInfo | Excel boolean localization de-DE | StringValue of logical formula cells | Validate localized TRUE FALSE output

using Aspose.Cells;
using System;
using System.Globalization;

// The example creates a new Workbook, sets its CultureInfo to de-DE, inserts =TRUE() and =FALSE() formulas in cells A1 and A2, calculates all formulas, reads the StringValue of each cell expecting the German terms "WAHR" and "FALSCH", validates the localized output, and saves the workbook as LocalizedBoolean.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Set the workbook's culture to German (de-DE) for localized Boolean strings
        workbook.Settings.CultureInfo = new CultureInfo("de-DE");

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Insert a logical formula that evaluates to TRUE
        Cell trueCell = sheet.Cells["A1"];
        trueCell.Formula = "=TRUE()";

        // Insert a logical formula that evaluates to FALSE
        Cell falseCell = sheet.Cells["A2"];
        falseCell.Formula = "=FALSE()";

        // Calculate all formulas in the workbook
        workbook.CalculateFormula();

        // Retrieve the displayed (localized) string values
        string trueDisplay = trueCell.StringValue;   // Expected "WAHR"
        string falseDisplay = falseCell.StringValue; // Expected "FALSCH"

        // Output the results
        Console.WriteLine($"A1 displays: {trueDisplay}");
        Console.WriteLine($"A2 displays: {falseDisplay}");

        // Validate the localization
        if (trueDisplay != "WAHR")
            Console.WriteLine("TRUE value is not localized correctly.");
        if (falseDisplay != "FALSCH")
            Console.WriteLine("FALSE value is not localized correctly.");

        // Save the workbook (optional)
        workbook.Save("LocalizedBoolean.xlsx");
    }
}
