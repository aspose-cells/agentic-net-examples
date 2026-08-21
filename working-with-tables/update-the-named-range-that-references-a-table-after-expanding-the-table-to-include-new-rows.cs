// Title: C# – Update a Named Range After Expanding an Aspose.Cells ListObject (Table)
// Description: Demonstrates how to create a workbook, define a ListObject, add rows, resize the table, refresh the named range with SetRefersTo, and use the range in a SUM formula using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | named range | ListObject resize | SetRefersTo | dynamic range | Excel table expansion | SUM formula | workbook automation | .NET Excel API
// Common Searches: Aspose.Cells update named range after table resize | C# resize ListObject and refresh named range | SetRefersTo table DataRange Aspose.Cells example | dynamic named range for expanding Excel table .NET | how to recalculate named range after adding rows Aspose
// Developer Intent: Synchronize a named range with the new size of a ListObject after the table has been expanded.
// Use Cases: Generate a sales report where the table grows daily and the named range must always reflect the current data for totals. | Create chart data sources that automatically adjust when rows are added to an Excel table. | Build automated workbook templates that add rows, resize tables, and keep dependent formulas accurate without manual updates.
// AI Prompts: Show C# code to update a named range after resizing an Aspose.Cells ListObject. | Explain how SetRefersTo works with a table's DataRange in Aspose.Cells for .NET. | Provide a step‑by‑step example of adding rows, resizing a table, and refreshing a named range for a SUM formula.

using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

// Demonstrates how to create a workbook, define a ListObject, add rows, resize the table, refresh the named range with SetRefersTo, and use the range in a SUM formula using Aspose.Cells for .NET.
class UpdateNamedRangeAfterTableResize
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate initial data for the table (5 rows, 2 columns)
            for (int row = 0; row < 5; row++)
            {
                cells[row, 0].PutValue("Item " + (row + 1));
                cells[row, 1].PutValue((row + 1) * 10);
            }

            // Create a table (ListObject) covering the initial data range A1:B5
            int tableIndex = sheet.ListObjects.Add(0, 0, 4, 1, true);
            ListObject table = sheet.ListObjects[tableIndex];
            table.DisplayName = "SalesTable";

            // Create a named range that refers to the table's data range
            int nameIndex = workbook.Worksheets.Names.Add("SalesTableRange");
            Name namedRange = workbook.Worksheets.Names[nameIndex];
            // Set the RefersTo formula to the current data range of the table (no leading '=')
            namedRange.SetRefersTo(table.DataRange.RefersTo, false, false);

            // Add additional rows to the worksheet (rows 6-10)
            for (int row = 5; row < 10; row++)
            {
                cells[row, 0].PutValue("Item " + (row + 1));
                cells[row, 1].PutValue((row + 1) * 10);
            }

            // Resize the table to include the new rows (now rows 0-9)
            table.Resize(0, 0, 9, 1, true);

            // Update the named range to point to the expanded table range
            namedRange.SetRefersTo(table.DataRange.RefersTo, false, false);

            // Demonstrate that the named range works in a formula
            cells["D1"].Formula = "=SUM(SalesTableRange)";
            workbook.CalculateFormula();

            // Save the workbook
            string outputPath = "UpdatedNamedRange.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
