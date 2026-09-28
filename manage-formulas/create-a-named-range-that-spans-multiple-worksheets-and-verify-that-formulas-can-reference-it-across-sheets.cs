// Title: How to define a named range that spans multiple worksheets and use it in a SUM formula with Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code with Aspose.Cells that creates a named range covering A1:A2 on two worksheets and inserts a =SUM(namedRange) formula on a third worksheet. | Show how to trigger formula calculation and read the resulting value of a multi‑sheet named range in a .NET workbook using Aspose.Cells.
// Common Searches: how to create a named range that includes cells from multiple sheets using Aspose.Cells C# | using Aspose.Cells C# to sum values from a multi‑sheet named range | Aspose.Cells calculate formulas after adding a cross‑sheet named range | verify result of SUM of named range spanning several worksheets in Aspose.Cells | save workbook with multi‑sheet named range to XLSX using Aspose.Cells .NET
// Tags: Aspose.Cells define multi‑sheet named range | Aspose.Cells SUM formula with named range | C# Aspose.Cells calculate workbook formulas | Aspose.Cells reference named range across worksheets | Aspose.Cells save workbook to XLSX

using System;
using System.IO;
using Aspose.Cells;

// The program creates a workbook with three sheets, defines a named range that includes A1:A2 on Sheet1 and Sheet2, uses that range in a =SUM formula on Sheet3, calculates the formulas (resulting in 100), outputs the sum, and saves the file as MultiSheetNamedRange.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Rename the default sheet and add a second sheet
            Worksheet sheet1 = workbook.Worksheets[0];
            sheet1.Name = "Sheet1";
            Worksheet sheet2 = workbook.Worksheets.Add("Sheet2");

            // Populate some numeric data in both sheets
            sheet1.Cells["A1"].PutValue(10);
            sheet1.Cells["A2"].PutValue(20);
            sheet2.Cells["A1"].PutValue(30);
            sheet2.Cells["A2"].PutValue(40);

            // Create a named range that spans A1:A2 on both Sheet1 and Sheet2
            // Add returns the index of the new name; retrieve the Name object via the index
            int rangeIndex = workbook.Worksheets.Names.Add("MyMultiSheetRange");
            Name multiSheetRange = workbook.Worksheets.Names[rangeIndex];
            multiSheetRange.RefersTo = "Sheet1!$A$1:$A$2,Sheet2!$A$1:$A$2";

            // Add a third sheet to test referencing the named range
            Worksheet sheet3 = workbook.Worksheets.Add("Sheet3");
            // Use the named range in a formula (sum all values across the two sheets)
            sheet3.Cells["A1"].Formula = "=SUM(MyMultiSheetRange)";

            // Calculate formulas in the workbook
            workbook.CalculateFormula();

            // Verify the result (expected 10+20+30+40 = 100)
            double sumResult = sheet3.Cells["A1"].DoubleValue;
            Console.WriteLine("Sum of MyMultiSheetRange = " + sumResult);

            // Save the workbook to a file
            string outputPath = "MultiSheetNamedRange.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
