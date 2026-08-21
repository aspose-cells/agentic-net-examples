// Title: Disable Excel table auto‑expand in C# using Aspose.Cells ConvertToRange
// Description: This example shows how to create a ListObject (Excel table) with Aspose.Cells for .NET, apply a style, and then call ConvertToRange() to lock the table's range. After conversion the table is removed, additional rows can be added below without expanding the original range, and the worksheet’s ListObjects collection becomes empty.
// Keywords: Aspose.Cells ConvertToRange | C# Excel table fixed range | disable table auto expand .NET | ListObject to range Aspose | prevent Excel table growth | Aspose.Cells ListObject example
// Common Searches: Aspose.Cells ConvertToRange C# | stop Excel table from expanding with Aspose | convert ListObject to range .NET | fixed table range Aspose.Cells | how to lock Excel table size programmatically
// Developer Intent: Transform a ListObject into a normal cell range so new rows do not automatically become part of the table.
// Use Cases: Create a styled table for initial data, then lock its size before appending summary rows. | Provide a template where the table dimensions must stay constant while users add extra entries below. | Maintain static formula references by fixing the table range prior to further data imports.
// AI Prompts: Generate C# code that builds an Excel table with Aspose.Cells, converts it to a fixed range, and adds rows without expanding the table. | Explain what happens to table styles, filters, and the ListObjects collection after ConvertToRange is called. | Show how to programmatically confirm that a ListObject has been removed and that ListObjects.Count equals zero.

using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AsposeCellsTableFixedRangeDemo
{
    // This example shows how to create a ListObject (Excel table) with Aspose.Cells for .NET, apply a style, and then call ConvertToRange() to lock the table's range. After conversion the table is removed, additional rows can be added below without expanding the original range, and the worksheet’s ListObjects collection becomes empty.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate initial data (including header row)
            sheet.Cells["A1"].PutValue("ID");
            sheet.Cells["B1"].PutValue("Name");
            for (int i = 2; i <= 5; i++) // rows 2‑5 contain data
            {
                sheet.Cells[i - 1, 0].PutValue(i - 1);               // ID
                sheet.Cells[i - 1, 1].PutValue($"Item {i - 1}");    // Name
            }

            // Add a ListObject (Excel table) covering the data range A1:B5
            int tableIndex = sheet.ListObjects.Add(0, 0, 4, 1, true);
            ListObject table = sheet.ListObjects[tableIndex];
            table.TableStyleType = TableStyleType.TableStyleMedium2;

            // ------------------------------------------------------------
            // Disable auto‑expand: convert the table back to a normal range.
            // After conversion the table object is removed and its range
            // remains fixed even if more rows are added below.
            // ------------------------------------------------------------
            table.ConvertToRange();

            // Add extra rows below the original range – the former table will NOT expand.
            for (int i = 6; i <= 9; i++) // rows 6‑9 are new data
            {
                sheet.Cells[i - 1, 0].PutValue(i - 1);
                sheet.Cells[i - 1, 1].PutValue($"NewItem {i - 1}");
            }

            // Verify that the ListObjects collection is now empty
            Console.WriteLine("ListObjects count after conversion: " + sheet.ListObjects.Count);

            // Save the workbook
            workbook.Save("TableFixedRange.xlsx");
        }
    }
}
