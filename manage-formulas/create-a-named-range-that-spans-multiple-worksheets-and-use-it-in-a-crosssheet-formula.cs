// Title: Aspose.Cells for .NET: Create a Multi‑Sheet Named Range and Use It in a SUM Formula (C#)
// Description: Demonstrates how to build a workbook with two worksheets, define a named range that spans Sheet1!A1:A3 and Sheet2!B1:B3 using the RefersTo property, apply the range in a cross‑sheet =SUM formula, calculate the result, retrieve each constituent range with GetRanges, and save the file as MultiSheetNamedRange.xlsx.
// Keywords: Aspose.Cells multi sheet named range | C# named range across worksheets | Aspose.Cells RefersTo property | cross sheet SUM formula Aspose.Cells | GetRanges Aspose.Cells | calculate formulas Aspose.Cells | Aspose.Cells example C#
// Common Searches: Aspose.Cells create named range on multiple sheets | How to use a multi‑sheet named range in a formula with Aspose.Cells | Retrieve areas of a named range spanning several worksheets Aspose.Cells | C# Aspose.Cells RefersTo syntax for non‑contiguous ranges | Sum values from different worksheets using a named range Aspose.Cells
// Developer Intent: Define a named range that includes cells from more than one worksheet and reference it in a formula.
// Use Cases: Combine ranges from Sheet1 and Sheet2 into a single named range and calculate their total with =SUM(MultiSheetRange). | Programmatically evaluate formulas after setting up the named range by calling workbook.CalculateFormula(). | Extract each sub‑range of a multi‑sheet named range via multiSheetRange.GetRanges() for reporting or further processing.
// AI Prompts: Generate C# code using Aspose.Cells to create a named range that covers Sheet1!A1:A5 and Sheet2!C1:C5 and apply it in an AVERAGE formula. | Show how to list all individual areas of a multi‑sheet named range with Aspose.Cells. | Explain the correct RefersTo string format for a named range that includes non‑contiguous ranges on different worksheets.

using System;
using Aspose.Cells;

// Demonstrates how to build a workbook with two worksheets, define a named range that spans Sheet1!A1:A3 and Sheet2!B1:B3 using the RefersTo property, apply the range in a cross‑sheet =SUM formula, calculate the result, retrieve each constituent range with GetRanges, and save the file as MultiSheetNamedRange.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet and rename it
            Worksheet sheet1 = workbook.Worksheets[0];
            sheet1.Name = "Sheet1";

            // Add a second worksheet
            Worksheet sheet2 = workbook.Worksheets.Add("Sheet2");

            // Populate data in Sheet1 (A1:A3)
            sheet1.Cells["A1"].PutValue(1);
            sheet1.Cells["A2"].PutValue(2);
            sheet1.Cells["A3"].PutValue(3);

            // Populate data in Sheet2 (B1:B3)
            sheet2.Cells["B1"].PutValue(10);
            sheet2.Cells["B2"].PutValue(20);
            sheet2.Cells["B3"].PutValue(30);

            // Create a named range that spans both worksheets
            int nameIndex = workbook.Worksheets.Names.Add("MultiSheetRange");
            Name multiSheetRange = workbook.Worksheets.Names[nameIndex];
            // Union of two areas: Sheet1!A1:A3 and Sheet2!B1:B3
            multiSheetRange.RefersTo = "=Sheet1!$A$1:$A$3,Sheet2!$B$1:$B$3";

            // Use the named range in a cross‑sheet formula (sum of all cells in the range)
            sheet1.Cells["C1"].Formula = "=SUM(MultiSheetRange)";

            // Calculate formulas so that C1 gets the result
            workbook.CalculateFormula();

            // Retrieve individual ranges that compose the named range
            Aspose.Cells.Range[] ranges = multiSheetRange.GetRanges();
            Console.WriteLine($"Named range consists of {ranges.Length} areas:");
            foreach (Aspose.Cells.Range r in ranges)
            {
                Console.WriteLine($"{r.Worksheet.Name}!{r.Address}");
            }

            // Display the result of the formula
            Console.WriteLine($"Sum result in C1: {sheet1.Cells["C1"].Value}");

            // Save the workbook
            workbook.Save("MultiSheetNamedRange.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
