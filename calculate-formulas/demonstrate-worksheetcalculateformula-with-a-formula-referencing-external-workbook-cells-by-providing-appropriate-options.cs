// Title: Calculate formulas that reference an external workbook using Worksheet.CalculateFormula and LinkedDataSources in Aspose.Cells for .NET
// AI Prompts: Show C# code that creates a secondary workbook, adds it to CalculationOptions.LinkedDataSources, and calls Worksheet.CalculateFormula to evaluate a direct external cell reference. | Demonstrate evaluating an INDIRECT formula that points to a cell in another workbook by configuring CalculationOptions with linked data sources in Aspose.Cells.
// Common Searches: asp.net aspose.cells calculate formula that points to another workbook | using CalculationOptions.LinkedDataSources to evaluate external cell reference in C# | Worksheet.CalculateFormula example with external workbook and INDIRECT function | how to compute external workbook formulas without opening the source file in Aspose.Cells
// Tags: Worksheet.CalculateFormula external workbook reference | Aspose.Cells linked data sources formula evaluation | calculate INDIRECT function with external workbook in .NET | external workbook cell value calculation using CalculationOptions

using System;
using Aspose.Cells;

// The example creates an external workbook containing a value in cell A2, sets up a main workbook with formulas that reference that external cell directly and via the INDIRECT function, configures CalculationOptions.LinkedDataSources to include the external workbook, and uses Worksheet.CalculateFormula to compute both formulas, outputting the results.
class Program
{
    static void Main()
    {
        // Create an external workbook that will serve as the data source
        Workbook externalWb = new Workbook();
        Worksheet externalSheet = externalWb.Worksheets[0];
        externalSheet.Name = "Sheet1";
        // Put a sample value in cell A2 of the external workbook
        externalSheet.Cells["A2"].PutValue(12345);

        // Create the main workbook where formulas will reference the external workbook
        Workbook mainWb = new Workbook();
        Worksheet mainSheet = mainWb.Worksheets[0];

        // Set formulas that reference the external workbook
        // Direct external reference
        mainSheet.Cells["A1"].Formula = "=[External.xlsx]Sheet1!$A$2";
        // INDIRECT function referencing the external workbook
        mainSheet.Cells["A2"].Formula = "=INDIRECT(\"[External.xlsx]Sheet1!$A$2\")";

        // Prepare calculation options and link the external workbook as a data source
        CalculationOptions calcOptions = new CalculationOptions
        {
            LinkedDataSources = new Workbook[] { externalWb }
        };

        // Calculate the formulas using Worksheet.CalculateFormula with the provided options
        object directResult = mainSheet.CalculateFormula("=[External.xlsx]Sheet1!$A$2", calcOptions);
        object indirectResult = mainSheet.CalculateFormula("=INDIRECT(\"[External.xlsx]Sheet1!$A$2\")", calcOptions);

        // Output the calculated results
        Console.WriteLine("Direct external reference result: " + directResult);
        Console.WriteLine("INDIRECT external reference result: " + indirectResult);
    }
}
