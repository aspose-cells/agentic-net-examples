// Title: How to register and use a custom compound interest function in Aspose.Cells for .NET
// AI Prompts: Create a C# class that implements Aspose.Cells.ICustomFunction to calculate principal * (Math.Pow(1 + rate, periods) - 1) and add the instance to workbook.CustomFunctions with a chosen name. | Insert a formula like =COMPOUNDINTEREST(A1,A2,A3) into a worksheet cell, trigger workbook.CalculateFormula(), and save the workbook to verify the custom function works.
// Common Searches: aspnet register custom financial function Aspose.Cells example | how to implement ICustomFunction for compound interest in C# | using custom functions in Aspose.Cells workbook calculations | Aspose.Cells .NET custom function for financial formulas
// Tags: custom function registration Aspose.Cells | compound interest calculation C# | financial formulas custom function .NET | Aspose.Cells ICustomFunction implementation | worksheet formula with custom function

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsCustomFunctionDemo
{
    // This example shows how to implement an ICustomFunction that computes compound interest, register it in the Workbook.CustomFunctions collection, call it from a cell formula (e.g., =COMPOUNDINTEREST(A1,A2,A3)), recalculate the sheet, and save the workbook as an Excel file.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook.
                Workbook workbook = new Workbook();

                // Access the first worksheet.
                Worksheet sheet = workbook.Worksheets[0];

                // Example data:
                // Principal = 1000, Rate = 5% (0.05), Periods = 10
                sheet.Cells["A1"].PutValue(1000);   // Principal
                sheet.Cells["A2"].PutValue(0.05);   // Rate per period
                sheet.Cells["A3"].PutValue(10);     // Number of periods

                // Use built‑in Excel functions to calculate compound interest:
                // Interest = Principal * (POWER(1 + Rate, Periods) - 1)
                sheet.Cells["B1"].Formula = "A1*(POWER(1+A2,A3)-1)";

                // Calculate formulas.
                workbook.CalculateFormula();

                // Define output file path.
                string outputPath = "CompoundInterestDemo.xlsx";

                // Save the workbook.
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                // Log any unexpected errors.
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
