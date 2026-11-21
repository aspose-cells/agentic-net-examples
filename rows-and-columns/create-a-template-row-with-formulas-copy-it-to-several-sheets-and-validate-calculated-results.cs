// Title: Copy a template row containing a formula to several worksheets and validate the results with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# copy a row with formulas to another worksheet | how to recalculate formulas after copying rows in Aspose.Cells .NET | verify cell values after copying template row across multiple sheets using Aspose.Cells | using CopyRows method to duplicate a formula row in Excel workbook with Aspose.Cells | C# example for validating calculated results in copied rows with Aspose.Cells
// Tags: CopyRows method with formulas Aspose.Cells | template row duplication across worksheets .NET | recalculate workbook formulas Aspose.Cells | validate calculated cell values C# Aspose.Cells | Excel row copy preserving formulas Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// This C# example creates a workbook, defines a template row with static values and a formula (C1 = A1 + B1), copies that row to additional worksheets using the CopyRows method, recalculates all formulas, checks that each sheet's C1 matches the expected sum, and saves the file as an XLSX document.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Add worksheets (the first one is created by default)
            Worksheet templateSheet = workbook.Worksheets[0];
            templateSheet.Name = "Template";
            Worksheet sheet2 = workbook.Worksheets.Add("Sheet2");
            Worksheet sheet3 = workbook.Worksheets.Add("Sheet3");

            // ----- Define the template row (row 0) -----
            // Static values
            templateSheet.Cells[0, 0].PutValue(5);   // A1 = 5
            templateSheet.Cells[0, 1].PutValue(10);  // B1 = 10
            // Formula: C1 = A1 + B1
            templateSheet.Cells[0, 2].Formula = "=A1+B1";

            // ----- Copy the template row to other sheets -----
            // Copy row 0 from templateSheet to row 0 of sheet2 and sheet3
            sheet2.Cells.CopyRows(templateSheet.Cells, 0, 0, 1);
            sheet3.Cells.CopyRows(templateSheet.Cells, 0, 0, 1);

            // Recalculate all formulas in the workbook
            workbook.CalculateFormula();

            // ----- Validate calculated results -----
            double result1 = templateSheet.Cells[0, 2].DoubleValue; // C1 in Template
            double result2 = sheet2.Cells[0, 2].DoubleValue;        // C1 in Sheet2
            double result3 = sheet3.Cells[0, 2].DoubleValue;        // C1 in Sheet3

            const double expected = 15.0; // 5 + 10

            Console.WriteLine($"Template!C1 = {result1} (expected {expected})");
            Console.WriteLine($"Sheet2!C1    = {result2} (expected {expected})");
            Console.WriteLine($"Sheet3!C1    = {result3} (expected {expected})");

            if (Math.Abs(result1 - expected) < 1e-9 &&
                Math.Abs(result2 - expected) < 1e-9 &&
                Math.Abs(result3 - expected) < 1e-9)
            {
                Console.WriteLine("All calculated results are correct.");
            }
            else
            {
                Console.WriteLine("Validation failed.");
            }

            // Save the workbook
            string outputPath = "TemplateCopyResult.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
