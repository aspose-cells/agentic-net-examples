// Title: How to disable auto‑expand for an Aspose.Cells ListObject to keep the table range fixed in C#
// Common Searches: how to keep an Aspose.Cells ListObject range from expanding when adding new rows | Aspose.Cells C# prevent table auto‑expand after inserting data | disable auto expand for Excel table using Aspose.Cells API | fixed table range in Aspose.Cells without ShowAutoExpand property | Aspose.Cells ListObject static range example in .NET
// Tags: Aspose.Cells ListObject fixed range | disable auto expand ListObject C# | Aspose.Cells table static range | prevent Excel table auto expansion .NET | ListObject range control Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AsposeCellsExample
{
    // The example creates a new workbook, defines a ListObject (Excel table) with a specific range, and notes that recent Aspose.Cells versions do not expose a ShowAutoExpand property, so the table range stays fixed unless rows are explicitly added to the table. The workbook is saved as TableWithFixedRange.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Fill sample data
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

                // Define table range (zero‑based indices)
                int firstRow = 0;      // Row 1
                int firstColumn = 0;   // Column A
                int totalRows = 4;     // Header + 3 data rows
                int totalColumns = 3;  // Columns A‑C

                // Add a ListObject (Excel table) to the worksheet
                int tableIndex = sheet.ListObjects.Add(firstRow, firstColumn,
                    firstRow + totalRows, firstColumn + totalColumns, true);
                ListObject table = sheet.ListObjects[tableIndex];

                // Note: Aspose.Cells does not expose a ShowAutoExpand property in recent versions.
                // The table range will remain fixed unless rows are explicitly added to the table.

                // Save the workbook
                string outputPath = "TableWithFixedRange.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
