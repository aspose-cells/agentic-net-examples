// Title: Disable Auto‑Expand of an Excel Table (ListObject) in Aspose.Cells for .NET
// Description: This C# example shows how to prevent an Aspose.Cells ListObject from automatically expanding. The workbook creates a table on A1:B3, calls ConvertToRange() to turn the table into a normal range, verifies that no ListObjects remain, adds extra rows that stay outside the former table, and saves the file.
// Keywords: Aspose.Cells | Aspose.Cells for .NET | C# | .NET | Excel table | ListObject | disable auto expand | ConvertToRange | fixed table range | prevent table expansion | GitHub sample | code example
// Common Searches: how to stop Aspose.Cells ListObject from expanding | convert Excel table to range Aspose.Cells C# | keep table range static in Aspose.Cells | disable auto‑expand of Excel table using Aspose.Cells | Aspose.Cells example for fixed table size
// Developer Intent: Keep an Excel table’s range static by converting the ListObject to a normal range.
// Use Cases: Create a table for initial data, then lock its size to prevent further auto‑expansion. | Add rows after conversion without them being incorporated into the original table. | Generate reports where the table layout must remain unchanged despite additional data.
// AI Prompts: Show C# code to disable auto‑expand of a ListObject in Aspose.Cells. | Explain how ConvertToRange works and how to verify that no ListObjects exist. | Provide a step‑by‑step guide to add rows below a converted range without affecting the original table.

using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AsposeCellsTableAutoExpandDemo
{
    // This C# example shows how to prevent an Aspose.Cells ListObject from automatically expanding. The workbook creates a table on A1:B3, calls ConvertToRange() to turn the table into a normal range, verifies that no ListObjects remain, adds extra rows that stay outside the former table, and saves the file.
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
            sheet.Cells["A2"].PutValue(1);
            sheet.Cells["B2"].PutValue("John");
            sheet.Cells["A3"].PutValue(2);
            sheet.Cells["B3"].PutValue("Mary");

            // Add a ListObject (Excel table) covering the data range A1:B3
            int tableIndex = sheet.ListObjects.Add("A1", "B3", true);
            ListObject table = sheet.ListObjects[tableIndex];

            // At this point the table will auto‑expand when new rows are added below it.
            // To keep the range fixed, convert the table back to a normal range.
            table.ConvertToRange();

            // Verify that the ListObject no longer exists
            Console.WriteLine("ListObjects count after conversion: " + sheet.ListObjects.Count); // Should be 0

            // Add more rows below the original range; they will NOT be part of the former table
            sheet.Cells["A4"].PutValue(3);
            sheet.Cells["B4"].PutValue("Alice");
            sheet.Cells["A5"].PutValue(4);
            sheet.Cells["B5"].PutValue("Bob");

            // Save the workbook
            workbook.Save("TableAutoExpandDisabled.xlsx");
        }
    }
}
