// Title: How to calculate a SUM formula that references an external workbook using Worksheet.CalculateFormula and custom CalculationOptions in Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# calculate formula that pulls data from another Excel file | using CalculationOptions to enable recursive calculation for external workbook references in Aspose.Cells | C# example of SUM across external workbook with Worksheet.CalculateFormula | how to ignore errors while evaluating external references in Aspose.Cells | set up external workbook reference syntax in Aspose.Cells formula
// Tags: external workbook reference handling Aspose.Cells | Worksheet.CalculateFormula with custom options | recursive formula evaluation Aspose.Cells | ignore calculation errors .NET | SUM across multiple Excel workbooks C#

using System;
using System.IO;
using Aspose.Cells;

// The example creates an external workbook, fills cells A1:A5 with numbers, saves it, then builds a main workbook that contains a SUM formula referencing the external sheet. It configures CalculationOptions (Recursive = true, IgnoreError = true), runs Workbook.CalculateFormula, reads the computed value from cell B2, prints the result, and saves the main workbook.
class Program
{
    static void Main()
    {
        try
        {
            // Prepare a temporary folder for the workbooks
            string folderPath = Path.Combine(Directory.GetCurrentDirectory(), "TempData");
            Directory.CreateDirectory(folderPath);

            // -----------------------------
            // Create the external workbook
            // -----------------------------
            string externalFile = Path.Combine(folderPath, "ExternalWorkbook.xlsx");
            Workbook externalWb = new Workbook();
            Worksheet extSheet = externalWb.Worksheets[0];
            extSheet.Name = "Sheet1";

            // Fill cells A1:A5 with sample numbers
            for (int i = 0; i < 5; i++)
            {
                extSheet.Cells[i, 0].PutValue(i + 1); // A1..A5 = 1,2,3,4,5
            }

            // Save the external workbook
            externalWb.Save(externalFile);

            // -----------------------------
            // Create the main workbook
            // -----------------------------
            Workbook mainWb = new Workbook();
            Worksheet mainSheet = mainWb.Worksheets[0];
            mainSheet.Name = "Main";

            // Set a formula that references the external workbook
            // Note the external reference syntax: '[FileName]SheetName'!Range
            string formula = $"=SUM('[ExternalWorkbook.xlsx]Sheet1'!A1:A5)";
            mainSheet.Cells["B2"].Formula = formula;

            // Prepare calculation options (e.g., allow recursive calculation)
            CalculationOptions calcOptions = new CalculationOptions
            {
                Recursive = true,          // Allow formulas that depend on other formulas
                IgnoreError = true         // Ignore errors during calculation
            };

            // Calculate all formulas in the workbook using the specified options
            mainWb.CalculateFormula(calcOptions);

            // Retrieve and display the calculated result
            double result = mainSheet.Cells["B2"].DoubleValue;
            Console.WriteLine($"Result of external reference formula: {result}");

            // (Optional) Save the main workbook to verify the calculated value is stored
            string mainFile = Path.Combine(folderPath, "MainWorkbook.xlsx");
            mainWb.Save(mainFile);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
