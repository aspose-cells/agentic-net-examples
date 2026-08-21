// Title: Define a multi‑sheet named range and sum its values with Aspose.Cells for .NET (C#)
// Description: This example creates a workbook, adds two worksheets with numeric data, defines a named range that spans both sheets, inserts a SUM formula on a third sheet that references the named range, calculates the formula, outputs the result, and saves the file as NamedRangeConsolidation.xlsx.
// Keywords: Aspose.Cells named range multiple worksheets | C# multi‑area named range | sum across sheets Aspose.Cells | consolidate data Aspose.Cells .NET | named range formula Aspose.Cells
// Common Searches: Aspose.Cells create named range across sheets | C# sum values from multiple worksheets using named range | multi‑sheet named range example Aspose.Cells | how to use SUM with a named range in Aspose.Cells
// Developer Intent: Create a named range that includes cells from several worksheets and use it in a formula to calculate a consolidated total.
// Use Cases: Combine monthly sales numbers from department sheets into a single total on a summary sheet. | Aggregate inventory counts from regional worksheets without writing individual cell references. | Apply other aggregate functions (AVERAGE, COUNT, MAX) to the same multi‑sheet range for dynamic reporting.
// AI Prompts: Generate C# code that defines a named range covering A1:A10 on SheetA and SheetB, then uses the AVERAGE function on a third sheet. | Show how to create a multi‑area named range in Aspose.Cells and apply a COUNT formula to count non‑empty cells across the referenced sheets. | Provide an example that updates the named range to include additional worksheets before recalculating a SUM on a summary sheet.

using System;
using Aspose.Cells;

namespace AsposeCellsNamedRangeConsolidation
{
    // This example creates a workbook, adds two worksheets with numeric data, defines a named range that spans both sheets, inserts a SUM formula on a third sheet that references the named range, calculates the formula, outputs the result, and saves the file as NamedRangeConsolidation.xlsx.
    class Program
    {
        static void Main()
        {
            // ---------- Create a new workbook ----------
            Workbook workbook = new Workbook();

            // ---------- Add two worksheets and populate them ----------
            Worksheet sheet1 = workbook.Worksheets[0];
            sheet1.Name = "Sheet1";
            sheet1.Cells["A1"].PutValue(10);
            sheet1.Cells["A2"].PutValue(20);
            sheet1.Cells["A3"].PutValue(30);

            Worksheet sheet2 = workbook.Worksheets.Add("Sheet2");
            sheet2.Cells["A1"].PutValue(5);
            sheet2.Cells["A2"].PutValue(15);
            sheet2.Cells["A3"].PutValue(25);

            // ---------- Define a named range that spans both worksheets ----------
            // The RefersTo string can contain multiple areas separated by commas.
            // Note the leading '=' required by Aspose.Cells.
            int nameIndex = workbook.Worksheets.Names.Add("ConsolidatedData");
            Name namedRange = workbook.Worksheets.Names[nameIndex];
            namedRange.RefersTo = "=Sheet1!$A$1:$A$3,Sheet2!$A$1:$A$3";

            // ---------- Add a third worksheet to display the consolidation result ----------
            Worksheet sheet3 = workbook.Worksheets.Add("Result");
            // Use the named range in a formula to sum all values across the two sheets.
            sheet3.Cells["A1"].Formula = "=SUM(ConsolidatedData)";

            // ---------- Calculate formulas ----------
            workbook.CalculateFormula();

            // ---------- Output the consolidated sum to console ----------
            Console.WriteLine("Consolidated Sum (Sheet1 + Sheet2): " + sheet3.Cells["A1"].Value);

            // ---------- Save the workbook ----------
            workbook.Save("NamedRangeConsolidation.xlsx");
        }
    }
}
