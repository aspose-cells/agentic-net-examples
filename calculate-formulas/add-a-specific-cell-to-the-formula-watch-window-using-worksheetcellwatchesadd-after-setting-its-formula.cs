// Title: Add a cell to the Formula Watch Window after assigning a SUM formula using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that writes values to cells B1‑B5, sets a SUM formula in A1, adds A1 to the Formula Watch Window with Worksheet.CellWatches.Add, prints the watch count, and saves the workbook. | Show how to retrieve and display the number of cells currently tracked in the Formula Watch Window after adding a watch in Aspose.Cells. | Create an example that demonstrates monitoring a formula’s result by adding the target cell to the watch list and then exporting the workbook to an .xlsx file.
// Common Searches: how to use Worksheet.CellWatches.Add to monitor a formula in Aspose.Cells C# | Aspose.Cells example adding a cell to the Formula Watch Window after setting a formula | C# code to display count of formula watches in Aspose.Cells workbook | track SUM calculation with Formula Watch Window using Aspose.Cells for .NET
// Tags: Aspose.Cells Worksheet.CellWatches.Add example | add cell to Formula Watch Window C# | monitor SUM formula Aspose.Cells | display formula watch count Aspose.Cells | save workbook after adding formula watch Aspose.Cells

using System;
using Aspose.Cells;

namespace AsposeCellsWatchDemo
{
    // Demonstrates populating cells B1‑B5, setting a SUM formula in A1, adding A1 to the Formula Watch Window via Worksheet.CellWatches.Add, printing the watch count, and saving the workbook as an .xlsx file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate cells that will be used in the formula
            sheet.Cells["B1"].PutValue(10);
            sheet.Cells["B2"].PutValue(20);
            sheet.Cells["B3"].PutValue(30);
            sheet.Cells["B4"].PutValue(40);
            sheet.Cells["B5"].PutValue(50);

            // Set a formula in cell A1
            sheet.Cells["A1"].Formula = "=SUM(B1:B5)";

            // Add cell A1 to the Formula Watch Window (using CellWatches.Add)
            sheet.CellWatches.Add("A1");

            // Optional: display the number of watches added
            Console.WriteLine("Cell watches count: " + sheet.CellWatches.Count);

            // Save the workbook (lifecycle rule: save)
            workbook.Save("FormulaWatchDemo.xlsx");
        }
    }
}
