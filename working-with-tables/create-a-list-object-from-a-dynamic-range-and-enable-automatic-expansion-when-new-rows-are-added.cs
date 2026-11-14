// Title: Create and resize an Aspose.Cells ListObject (Excel table) from a dynamic range with automatic expansion in C#
// Common Searches: how to add a ListObject to a worksheet from a dynamic range using Aspose.Cells C# | programmatically expand an Aspose.Cells Excel table when new rows are inserted | resize Aspose.Cells ListObject after adding rows in .NET | apply table style and autofit columns for a ListObject created with Aspose.Cells | save workbook with a dynamically growing table using Aspose.Cells for .NET
// Tags: Aspose.Cells ListObject creation from range | dynamic Excel table resizing Aspose.Cells | auto‑expand ListObject rows C# | apply table style Aspose.Cells | auto‑fit columns Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace DynamicTableExample
{
    // The example creates a new workbook, fills cells A1:C4 with sample data, adds a ListObject (Excel table) covering that range, sets a display name and style, inserts an additional row, resizes the ListObject to include the new row, auto‑fits the columns, and saves the file as DynamicTable.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                var workbook = new Workbook();

                // Access the first worksheet
                var sheet = workbook.Worksheets[0];

                // Populate sample data (dynamic range will start from A1)
                sheet.Cells["A1"].PutValue("ID");
                sheet.Cells["B1"].PutValue("Name");
                sheet.Cells["C1"].PutValue("Score");

                sheet.Cells["A2"].PutValue(1);
                sheet.Cells["B2"].PutValue("Alice");
                sheet.Cells["C2"].PutValue(85);

                sheet.Cells["A3"].PutValue(2);
                sheet.Cells["B3"].PutValue("Bob");
                sheet.Cells["C3"].PutValue(92);

                sheet.Cells["A4"].PutValue(3);
                sheet.Cells["B4"].PutValue("Charlie");
                sheet.Cells["C4"].PutValue(78);

                // Define the initial range for the list object (A1:C4)
                int firstRow = 0;   // zero‑based index for row 1
                int firstCol = 0;   // zero‑based index for column A
                int lastRow = 3;    // row 4
                int lastCol = 2;    // column C

                // Add a ListObject (Excel Table) based on the range
                int listObjectIndex = sheet.ListObjects.Add(firstRow, firstCol, lastRow, lastCol, true);
                var listObject = sheet.ListObjects[listObjectIndex];

                // Set a meaningful display name for the table
                listObject.DisplayName = "DynamicTable";

                // Apply a table style
                listObject.TableStyleType = TableStyleType.TableStyleMedium9;

                // Example: Add a new row below the current table range
                int newRowIndex = lastRow + 1; // row 5 (zero‑based index 4)
                sheet.Cells[newRowIndex, firstCol].PutValue(4);
                sheet.Cells[newRowIndex, firstCol + 1].PutValue("Diana");
                sheet.Cells[newRowIndex, firstCol + 2].PutValue(88);

                // Resize the ListObject to include the new row
                listObject.Resize(firstRow, firstCol, newRowIndex, lastCol, true);

                // Auto‑fit columns for better appearance
                for (int col = firstCol; col <= lastCol; col++)
                {
                    sheet.AutoFitColumn(col);
                }

                // Save the workbook
                string outputPath = "DynamicTable.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
