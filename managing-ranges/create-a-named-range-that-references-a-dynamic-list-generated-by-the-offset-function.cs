// Title: Define a Dynamic Named Range with OFFSET & COUNTA in Aspose.Cells for .NET
// Description: Shows how to build a workbook, populate column A, add a named range that leverages OFFSET and COUNTA to expand automatically as rows are added, use the range in a SUM formula, recalculate the sheet, and save the file with Aspose.Cells for .NET.
// Keywords: Aspose.Cells | .NET | dynamic named range | OFFSET function | COUNTA | expandable Excel range | programmatic named range | Excel automation | data validation list | formula calculation
// Common Searches: Aspose.Cells dynamic named range example | OFFSET function in Aspose.Cells .NET | use COUNTA to create expanding range Aspose | how to add named range programmatically Aspose.Cells | sum values of a dynamic range using Aspose.Cells | Excel named range that grows with new rows
// Developer Intent: Create a named range that automatically adjusts its size with added data using OFFSET and COUNTA.
// Use Cases: Build a data‑validation list that grows when new items are appended. | Calculate totals for a dataset that changes size without updating formulas. | Reference a self‑adjusting range in charts, pivot tables, or conditional formatting.
// AI Prompts: Generate C# code to define a dynamic named range with OFFSET in Aspose.Cells and apply it in a SUM formula. | Explain how to refresh the range after inserting rows and recalculate dependent formulas. | Show how to retrieve the address of a dynamic named range at runtime using Aspose.Cells.

using System;
using Aspose.Cells;

// Shows how to build a workbook, populate column A, add a named range that leverages OFFSET and COUNTA to expand automatically as rows are added, use the range in a SUM formula, recalculate the sheet, and save the file with Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Create a new workbook (lifecycle rule: create)
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Name = "Data";

        // Populate column A with sample data.
        // The first cell is a header; the rest form a dynamic list.
        sheet.Cells["A1"].PutValue("Header");
        sheet.Cells["A2"].PutValue("Item1");
        sheet.Cells["A3"].PutValue("Item2");
        sheet.Cells["A4"].PutValue("Item3");
        // Additional rows can be added later; the named range will expand automatically.

        // Add a named range that uses the OFFSET function to refer to the dynamic list.
        // OFFSET(start, rows, cols, height, width)
        // Start at A2 (first data row), no row/col offset, height = COUNTA(A:A)-1 (exclude header), width = 1.
        int nameIndex = workbook.Worksheets.Names.Add("DynamicList");
        Name dynamicName = workbook.Worksheets.Names[nameIndex];
        dynamicName.RefersTo = "=OFFSET(Data!$A$2,0,0,COUNTA(Data!$A:$A)-1,1)";

        // Use the named range in a formula (e.g., SUM) to demonstrate it works.
        sheet.Cells["B1"].Formula = "=SUM(DynamicList)";

        // Calculate all formulas in the workbook.
        workbook.CalculateFormula();

        // Output the calculated result.
        Console.WriteLine("Sum of DynamicList: " + sheet.Cells["B1"].Value);

        // Save the workbook (lifecycle rule: save)
        workbook.Save("DynamicNamedRange.xlsx");
    }
}
