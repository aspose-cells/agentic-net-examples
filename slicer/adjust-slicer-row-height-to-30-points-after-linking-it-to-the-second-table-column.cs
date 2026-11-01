// Title: Set slicer row height to 30 points for a column‑linked slicer using Aspose.Cells in C#
// Common Searches: Aspose.Cells C# set slicer row height to 30 points | How to change slicer height after linking it to a table column with Aspose.Cells | Programmatically adjust slicer dimensions in an Excel workbook using Aspose.Cells .NET | C# example for setting slicer RowHeight property in Aspose.Cells | Adjust slicer size when adding it to a ListObject with Aspose.Cells
// Tags: Aspose.Cells slicer row height | C# link slicer to table column | Aspose.Cells adjust slicer dimensions | Excel slicer customization .NET | set slicer RowHeight property

using System;
using Aspose.Cells;
using Aspose.Cells.Tables;
using Aspose.Cells.Slicers;

namespace AsposeCellsExample
{
    // The example creates a workbook, adds a table with sample data, inserts a slicer linked to the table's second column at cell D1, sets the slicer's RowHeight to 30 points, and saves the workbook as output.xlsx.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet and rename it
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Name = "Data";

                // Populate sample data
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Value");
                sheet.Cells["A2"].PutValue("A");
                sheet.Cells["B2"].PutValue(10);
                sheet.Cells["A3"].PutValue("B");
                sheet.Cells["B3"].PutValue(20);
                sheet.Cells["A4"].PutValue("C");
                sheet.Cells["B4"].PutValue(30);

                // Define the range for the table (including header row)
                int firstRow = 0;          // zero‑based index for row 1
                int firstColumn = 0;       // zero‑based index for column A
                int totalRows = 4;         // header + 3 data rows
                int totalColumns = 2;      // two columns (A and B)

                // Add a ListObject (table) to the worksheet
                int tableIndex = sheet.ListObjects.Add(firstRow, firstColumn,
                    firstRow + totalRows - 1, firstColumn + totalColumns - 1, true);
                ListObject table = sheet.ListObjects[tableIndex];
                table.DisplayName = "MyTable";

                // Add a slicer linked to the second column of the table (zero‑based column index = 1)
                // Position the slicer at cell D1 using address string overload
                int slicerIndex = sheet.Slicers.Add(table, 1, "D1");
                Slicer slicer = sheet.Slicers[slicerIndex];

                // Adjust the slicer row height to 30 points
                slicer.RowHeight = 30;

                // Save the workbook
                workbook.Save("output.xlsx");
                Console.WriteLine("Workbook saved successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
