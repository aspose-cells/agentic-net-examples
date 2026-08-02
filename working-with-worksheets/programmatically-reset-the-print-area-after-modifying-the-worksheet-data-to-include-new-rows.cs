// Title: Aspose.Cells for .NET – Reset Worksheet Print Area After Adding Rows (C#)
// Description: This example creates a workbook, sets an initial print area, inserts new rows, adds data, and then updates PageSetup.PrintArea to cover the expanded range before saving. It demonstrates how to programmatically adjust the print area when worksheet content changes.
// Keywords: Aspose.Cells | .NET | C# | PrintArea | PageSetup.PrintArea | reset print area | update print area | insert rows | worksheet print area | dynamic print range
// Common Searches: Aspose.Cells reset print area C# | change print area after inserting rows Aspose.Cells | PageSetup.PrintArea dynamic range .NET | C# code to update worksheet print area | Aspose.Cells example print area
// Developer Intent: Modify the worksheet's print area so it includes rows added programmatically.
// Use Cases: Set a fixed print area, insert rows, then recalculate and assign a new PrintArea before saving. | Generate reports with a variable number of rows and automatically expand the print range to the used cells. | Create a template workbook, modify its data via code, and ensure the printed output captures all added rows.
// AI Prompts: Write C# code using Aspose.Cells to adjust the print area after rows are inserted. | Show how to determine the last used row in a worksheet and set PageSetup.PrintArea to the full data range. | Explain the steps to programmatically reset the print area when worksheet data size changes in an Aspose.Cells workbook.

using System;
using Aspose.Cells;

namespace ResetPrintAreaDemo
{
    // This example creates a workbook, sets an initial print area, inserts new rows, adds data, and then updates PageSetup.PrintArea to cover the expanded range before saving. It demonstrates how to programmatically adjust the print area when worksheet content changes.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate initial data (A1:B3)
            worksheet.Cells["A1"].PutValue("Header1");
            worksheet.Cells["B1"].PutValue("Header2");
            worksheet.Cells["A2"].PutValue("Item1");
            worksheet.Cells["B2"].PutValue(10);
            worksheet.Cells["A3"].PutValue("Item2");
            worksheet.Cells["B3"].PutValue(20);

            // Set the initial print area to cover the existing data
            worksheet.PageSetup.PrintArea = "A1:B3";

            // Insert two new rows after the existing data
            worksheet.Cells.InsertRows(3, 2); // Inserts rows at index 3 (zero‑based)

            // Add data to the newly inserted rows
            worksheet.Cells["A4"].PutValue("Item3");
            worksheet.Cells["B4"].PutValue(30);
            worksheet.Cells["A5"].PutValue("Item4");
            worksheet.Cells["B5"].PutValue(40);

            // Reset the print area to include the new rows (now rows 1‑5)
            worksheet.PageSetup.PrintArea = "A1:B5";

            // Save the workbook
            workbook.Save("ResetPrintAreaDemo.xlsx");

            Console.WriteLine("Workbook saved with updated print area.");
        }
    }
}
