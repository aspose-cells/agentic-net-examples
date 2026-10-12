// Title: How to make an Aspose.Cells ListObject (Excel table) auto‑expand when adding rows in C#
// AI Prompts: Write C# code using Aspose.Cells that creates a ListObject and demonstrates that adding a row below the table automatically expands its range. | Show how to apply a built‑in table style to a ListObject and then append data without manually resizing the table. | Provide a complete example that saves the workbook after the table has grown, handling any exceptions.
// Common Searches: Aspose.Cells C# auto expand Excel table when inserting new rows | C# add row to ListObject and have table range increase automatically | How to enable automatic table expansion in Aspose.Cells for .NET
// Tags: listobject dynamic range Aspose.Cells C# | append row to Excel table Aspose.Cells | default table style Aspose.Cells | save workbook after table growth C# | exception handling Aspose.Cells table

using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

// Creates a ListObject in a new workbook, applies a default table style, adds a row beneath the table, and relies on Aspose.Cells to automatically extend the table range before saving the file.
class TableAutoExpandExample
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate initial data for the table (A1:C3)
            sheet.Cells["A1"].PutValue("ID");
            sheet.Cells["B1"].PutValue("Name");
            sheet.Cells["C1"].PutValue("Score");

            sheet.Cells["A2"].PutValue(1);
            sheet.Cells["B2"].PutValue("Alice");
            sheet.Cells["C2"].PutValue(85);

            sheet.Cells["A3"].PutValue(2);
            sheet.Cells["B3"].PutValue("Bob");
            sheet.Cells["C3"].PutValue(92);

            // Define the range that will become the table
            CellArea tableArea = CellArea.CreateCellArea("A1", "C3");

            // Add a ListObject (Excel table) to the worksheet
            int tableIndex = sheet.ListObjects.Add(
                tableArea.StartRow,
                tableArea.StartColumn,
                tableArea.EndRow,
                tableArea.EndColumn,
                true);
            ListObject table = sheet.ListObjects[tableIndex];

            // OPTIONAL: Apply a built‑in table style for better visual appearance
            table.TableStyleType = TableStyleType.TableStyleMedium9;

            // Add new data directly below the current table range (row 4)
            sheet.Cells["A4"].PutValue(3);
            sheet.Cells["B4"].PutValue("Charlie");
            sheet.Cells["C4"].PutValue(78);

            // The table automatically expands to include the new row.
            // No manual resize is required.

            // Save the workbook to a file
            workbook.Save("TableAutoExpand.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
