// Title: Add a Formula Cell to the Formula Watch Window with Worksheet.CellWatches.Add (Aspose.Cells for .NET)
// Description: C# sample that creates a workbook, fills A1‑B2 with numbers, sets =SUM(A1:B2) in C3, adds C3 to the Formula Watch Window via Worksheet.CellWatches.Add(string), reads the watch’s address, row and column, and saves the workbook as FormulaWatchDemo.xlsx.
// Keywords: Aspose.Cells | Worksheet.CellWatches.Add | Formula Watch Window | C# formula monitoring | add cell watch .NET | debug spreadsheet formulas | track cell changes Aspose | monitor formula result | global developers | US developers
// Common Searches: How to add a cell to the Formula Watch Window in Aspose.Cells C# | Worksheet.CellWatches.Add example with formula | Retrieve CellWatch details after adding Aspose.Cells | Save workbook after adding formula watch Aspose.Cells | Aspose.Cells monitor formula cell during debugging
// Developer Intent: Add a specific cell that contains a formula to the Formula Watch Window and confirm its entry.
// Use Cases: Debug complex spreadsheets by watching the result of a critical calculation cell. | Ensure key summary cells recalculate correctly after data updates in large workbooks. | Audit formula locations by displaying watch details (cell name, row, column) before publishing.
// AI Prompts: Generate C# code to add multiple cells to the Formula Watch Window, iterate the CellWatch collection, and print each watch's details. | Show how to remove a CellWatch by address using Aspose.Cells and verify the removal. | Explain how to export watched cells, including formulas and current values, to a CSV file with Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsWatchDemo
{
    // C# sample that creates a workbook, fills A1‑B2 with numbers, sets =SUM(A1:B2) in C3, adds C3 to the Formula Watch Window via Worksheet.CellWatches.Add(string), reads the watch’s address, row and column, and saves the workbook as FormulaWatchDemo.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (uses the standard creation rule)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Set some sample values that the formula will use
            sheet.Cells["A1"].PutValue(10);
            sheet.Cells["A2"].PutValue(20);
            sheet.Cells["B1"].PutValue(30);
            sheet.Cells["B2"].PutValue(40);

            // Set a formula in cell C3
            Cell targetCell = sheet.Cells["C3"];
            targetCell.Formula = "=SUM(A1:B2)";

            // Add the cell to the Formula Watch Window using its address
            // This uses the Add(string) overload of CellWatchCollection
            int watchIndex = sheet.CellWatches.Add("C3");

            // Retrieve the added CellWatch to demonstrate that it was added
            CellWatch watch = sheet.CellWatches[watchIndex];
            Console.WriteLine($"Watch added at index {watchIndex}: Cell = {watch.CellName}, Row = {watch.Row}, Column = {watch.Column}");

            // Save the workbook (uses the standard save rule)
            workbook.Save("FormulaWatchDemo.xlsx");
        }
    }
}
