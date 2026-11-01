// Title: Calculate a formula that references an external workbook using Worksheet.CalculateFormula and CalculationOptions in Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# calculate formula that points to another workbook file | Using CalculationOptions to ignore errors when evaluating external references in Aspose.Cells | Worksheet.CalculateFormula example with external workbook reference in .NET | How to reference a cell from a different Excel file in Aspose.Cells formula calculation | C# code to evaluate external workbook formulas with Aspose.Cells
// Tags: external workbook formula calculation Aspose.Cells | Worksheet.CalculateFormula with CalculationOptions C# | ignore errors during formula evaluation Aspose.Cells | C# external cell reference Excel workbook Aspose.Cells | Aspose.Cells external reference formula example

using System;
using System.IO;
using Aspose.Cells;

// The sample creates a source workbook, writes a numeric value to A1, saves it, then builds a target workbook where cell B2 contains a formula that references the source workbook's A1 cell. It sets up CalculationOptions to ignore errors, runs Worksheet.CalculateFormula on the target workbook, reads the computed result from B2, prints it, and saves the target workbook.
class ExternalFormulaCalculationDemo
{
    static void Main()
    {
        try
        {
            // Create the source workbook and add a numeric value in A1
            Workbook sourceWb = new Workbook();
            Worksheet sourceWs = sourceWb.Worksheets[0];
            sourceWs.Name = "DataSheet";
            sourceWs.Cells["A1"].PutValue(123.45);

            // Save the source workbook (required for external reference)
            string sourcePath = "SourceWorkbook.xlsx";
            sourceWb.Save(sourcePath);

            // Verify the source file was saved successfully
            if (!File.Exists(sourcePath))
                throw new FileNotFoundException($"Source workbook not found: {sourcePath}");

            // Create the target workbook that will contain the external formula
            Workbook targetWb = new Workbook();
            Worksheet targetWs = targetWb.Worksheets[0];
            targetWs.Name = "CalcSheet";

            // Set a formula that references the external workbook's A1 cell
            targetWs.Cells["B2"].Formula = $"='[{sourcePath}]DataSheet'!A1";

            // Prepare calculation options (optional settings)
            CalculationOptions calcOptions = new CalculationOptions
            {
                // Example option: ignore errors during calculation
                IgnoreError = true
            };

            // Calculate all formulas in the target workbook using the options
            targetWb.CalculateFormula(calcOptions);

            // Retrieve and display the calculated result
            object result = targetWs.Cells["B2"].Value;
            Console.WriteLine($"Calculated value in B2: {result}");

            // Save the target workbook (optional, just to verify the result)
            string targetPath = "TargetWorkbook.xlsx";
            targetWb.Save(targetPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
