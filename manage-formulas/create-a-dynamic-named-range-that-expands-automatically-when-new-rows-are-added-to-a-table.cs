// Title: How to create a self‑expanding named range from a ListObject table using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# create a named range that automatically expands with new table rows | Define a dynamic named range using structured reference MyTable[#All] in .NET | Resize an Aspose.Cells ListObject and keep the named range up to date programmatically | C# example of auto‑growing named range linked to an Excel table | Aspose.Cells .NET add rows to a table and update the named range automatically
// Tags: Aspose.Cells dynamic named range | ListObject structured reference MyTable[#All] | C# table resize updates named range | auto‑expanding Excel named range .NET | Aspose.Cells workbook table manipulation

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AsposeCellsDynamicNamedRange
{
    // The example creates a new workbook, adds a ListObject table covering A1:C4, defines a named range 'MyDynamicRange' that references MyTable[#All] so it expands automatically, inserts an additional row, resizes the table to include the new data, and saves the file as DynamicNamedRange.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Get the first worksheet and name it
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Name = "Data";

                // Populate sample data (including header row)
                sheet.Cells["A1"].PutValue("ID");
                sheet.Cells["B1"].PutValue("Name");
                sheet.Cells["C1"].PutValue("Score");

                sheet.Cells["A2"].PutValue(1);
                sheet.Cells["B2"].PutValue("Alice");
                sheet.Cells["C2"].PutValue(85);

                sheet.Cells["A3"].PutValue(2);
                sheet.Cells["B3"].PutValue("Bob");
                sheet.Cells["C3"].PutValue(90);

                sheet.Cells["A4"].PutValue(3);
                sheet.Cells["B4"].PutValue("Charlie");
                sheet.Cells["C4"].PutValue(78);

                // Define a table (ListObject) that covers the data range A1:C4
                int firstRow = 0;   // zero‑based index for row 1
                int firstCol = 0;   // zero‑based index for column A
                int totalRows = 4;  // header + 3 data rows
                int totalCols = 3;  // columns A‑C

                // Add the table to the worksheet
                int tableIndex = sheet.ListObjects.Add(firstRow, firstCol,
                    firstRow + totalRows - 1, firstCol + totalCols - 1, true);
                ListObject table = sheet.ListObjects[tableIndex];
                table.DisplayName = "MyTable";   // Friendly name for the table
                table.ShowHeaderRow = true;

                // Create a dynamic named range that refers to the entire table.
                // The structured reference "MyTable[#All]" automatically expands when rows are added.
                int nameIndex = workbook.Worksheets.Names.Add("MyDynamicRange");
                Name dynamicRange = workbook.Worksheets.Names[nameIndex];
                dynamicRange.RefersTo = "=MyTable[#All]";

                // OPTIONAL: Demonstrate that adding a new row to the table expands the named range.
                // Determine the first empty row after the current data range.
                int newDataRow = table.DataRange.FirstRow + table.DataRange.RowCount;
                sheet.Cells[newDataRow, 0].PutValue(4);
                sheet.Cells[newDataRow, 1].PutValue("Diana");
                sheet.Cells[newDataRow, 2].PutValue(92);

                // Extend the table to include the new row.
                // New total rows = header row + existing data rows + the new row.
                int newTotalRows = table.DataRange.RowCount + 2; // +1 for header, +1 for new row
                // Resize requires a hasHeaders flag; the table has a header row.
                table.Resize(firstRow, firstCol, newTotalRows, totalCols, true);

                // Save the workbook
                string outputPath = "DynamicNamedRange.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
