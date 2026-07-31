// Title: Update a Named Range After Expanding an Aspose.Cells ListObject (C#)
// Description: Demonstrates how to create a workbook, define a ListObject (table) and a named range that points to the table's DataRange, add new rows, resize the table, and refresh the named range's RefersTo formula to cover the expanded range before saving the file.
// Keywords: Aspose.Cells C# update named range | resize ListObject Aspose.Cells | named range RefersTo table expansion | C# Aspose.Cells table resize example | dynamic named range Aspose.Cells
// Common Searches: Aspose.Cells update named range after table resize | C# resize ListObject and keep named range in sync | how to refresh named range when expanding a table Aspose.Cells | Aspose.Cells C# add rows and adjust named range | named range reference table data range Aspose.Cells
// Developer Intent: Refresh the named range so it points to the table’s new DataRange after the ListObject is resized.
// Use Cases: Add incoming data rows to a worksheet, resize the table, and keep formulas or charts aligned via an updated named range. | Generate reports that append summary rows and need a consistent named range for downstream pivot tables or external processing. | Synchronize named ranges with dynamic tables before exporting the workbook for consumption by other applications.
// AI Prompts: Write C# code using Aspose.Cells to add rows, resize an existing ListObject, and automatically update the associated named range. | Show an Aspose.Cells example that creates a named range from a table’s DataRange, expands the table, and refreshes the named range without recreating it. | Explain how to retrieve the updated address of a named range after a ListObject resize in Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AsposeCellsExamples
{
    // Demonstrates how to create a workbook, define a ListObject (table) and a named range that points to the table's DataRange, add new rows, resize the table, and refresh the named range's RefersTo formula to cover the expanded range before saving the file.
    public class UpdateNamedRangeAfterTableResize
    {
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "DataSheet";

            // Populate initial data for the table (5 rows, 2 columns)
            for (int i = 0; i < 5; i++)
            {
                sheet.Cells[i, 0].PutValue($"Item {i + 1}");
                sheet.Cells[i, 1].PutValue((i + 1) * 10);
            }

            // Create a ListObject (table) covering the initial data range A1:B5
            int tableIndex = sheet.ListObjects.Add(0, 0, 4, 1, true);
            ListObject table = sheet.ListObjects[tableIndex];
            table.DisplayName = "MyTable";

            // Create a named range that refers to the table's data range
            int nameIndex = workbook.Worksheets.Names.Add("MyTableRange");
            Name namedRange = workbook.Worksheets.Names[nameIndex];
            // Use the table's DataRange.Address to set the RefersTo formula
            namedRange.RefersTo = $"={sheet.Name}!{table.DataRange.Address}";

            // Add additional rows to the worksheet (rows 5 to 7)
            for (int i = 5; i < 8; i++)
            {
                sheet.Cells[i, 0].PutValue($"Item {i + 1}");
                sheet.Cells[i, 1].PutValue((i + 1) * 10);
            }

            // Resize the table to include the new rows (now rows 0 to 7)
            table.Resize(0, 0, 7, 1, true);

            // Update the named range to reference the expanded table range
            namedRange.RefersTo = $"={sheet.Name}!{table.DataRange.Address}";

            // Demonstrate that the named range now covers the new rows
            Aspose.Cells.Range range = namedRange.GetRange();
            Console.WriteLine($"Updated named range address: {range.Address}");
            Console.WriteLine($"Row count after resize: {range.RowCount}");

            // Save the workbook
            workbook.Save("UpdatedNamedRangeAfterTableResize.xlsx");
        }
    }
}
