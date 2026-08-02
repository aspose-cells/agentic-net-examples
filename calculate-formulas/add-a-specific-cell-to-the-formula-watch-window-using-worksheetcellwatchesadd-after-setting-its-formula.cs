// Title: Add a Cell to the Formula Watch Window with Worksheet.CellWatches.Add in Aspose.Cells for .NET (C#)
// Description: Creates a workbook, assigns a SUM formula to cell C3, adds that cell to the Formula Watch Window using Worksheet.CellWatches.Add, retrieves the watch entry to display its address, row and column, and saves the file so the watch data persists.
// Keywords: Aspose.Cells | Worksheet.CellWatches.Add | Formula Watch Window | C# example | track formula cell | CellWatch index | save workbook with watches | debug spreadsheet formulas
// Common Searches: add cell to formula watch window Aspose.Cells .NET | Worksheet.CellWatches.Add usage example | retrieve CellWatch details after adding | persist formula watch information in saved workbook | debug formulas with Aspose.Cells watch window
// Developer Intent: Add a formula‑containing cell to the Formula Watch Window, obtain its watch index and properties, and store the watch information in the workbook.
// Use Cases: Monitor key calculation cells during automated processing to verify correct evaluation. | Debug complex workbooks by watching specific cells and inspecting their runtime values. | Include watched cells in the saved file so reviewers can see which formulas were tracked.
// AI Prompts: Generate code that adds multiple cells to the Formula Watch Window and iterates through the CellWatches collection to list each address. | Show how to remove a CellWatch by address and update the watch list in an existing workbook. | Explain how to export all watched cells, including row and column indices, to a CSV file using Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsWatchDemo
{
    // Creates a workbook, assigns a SUM formula to cell C3, adds that cell to the Formula Watch Window using Worksheet.CellWatches.Add, retrieves the watch entry to display its address, row and column, and saves the file so the watch data persists.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Set a formula in cell C3 (row index 2, column index 2)
            Cell targetCell = sheet.Cells["C3"];
            targetCell.Formula = "=SUM(A1:B2)";

            // Add the cell to the Formula Watch Window using its address
            // The Add method returns the index of the watch item in the collection
            int watchIndex = sheet.CellWatches.Add("C3");

            // Retrieve the added CellWatch to verify its properties
            CellWatch watch = sheet.CellWatches[watchIndex];
            Console.WriteLine($"Watch added at index {watchIndex}: Cell = {watch.CellName}, Row = {watch.Row}, Column = {watch.Column}");

            // Save the workbook (the watch information is stored in the file)
            workbook.Save("FormulaWatchDemo.xlsx");
        }
    }
}
