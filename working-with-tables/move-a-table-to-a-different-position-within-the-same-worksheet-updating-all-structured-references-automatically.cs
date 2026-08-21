// Title: Move an Excel Table Within a Worksheet and Auto‑Update Structured References – Aspose.Cells for .NET (C#)
// Description: C# example that creates a workbook, defines a ListObject (table) on A1:B4, then uses Worksheet.Cells.MoveRange to relocate the table to row 10, column C. All structured references are refreshed automatically and the file is saved as MovedTable.xlsx.
// Keywords: Aspose.Cells | C# | MoveRange | ListObject | Excel table relocation | structured references | worksheet MoveRange example | programmatic table move | Excel API | Aspose.Cells sample code
// Common Searches: how to move an Excel table with Aspose.Cells C# | update structured references after moving a ListObject | Worksheet.Cells.MoveRange table example | relocate Excel table in same worksheet using .NET | Aspose.Cells move table without breaking formulas
// Developer Intent: Shift a ListObject to a new range on the same worksheet while keeping all structured references and formulas intact.
// Use Cases: Re‑position a data table after inserting rows above it without corrupting table formulas. | Organize a report layout by moving tables to designated sections while preserving calculations. | Automate dynamic worksheet designs where tables need to fit page‑size constraints or printing areas.
// AI Prompts: Write C# code that moves a ListObject to a different location in the same worksheet using Aspose.Cells, ensuring structured references are updated. | Show how to obtain a table's CellArea and apply Worksheet.Cells.MoveRange to relocate it without breaking formulas. | Explain the behavior of MoveRange for tables in Aspose.Cells and note any limitations or special considerations.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;

// C# example that creates a workbook, defines a ListObject (table) on A1:B4, then uses Worksheet.Cells.MoveRange to relocate the table to row 10, column C. All structured references are refreshed automatically and the file is saved as MovedTable.xlsx.
class MoveTableDemo
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data
            worksheet.Cells["A1"].PutValue("ID");
            worksheet.Cells["B1"].PutValue("Name");
            worksheet.Cells["A2"].PutValue(1);
            worksheet.Cells["B2"].PutValue("John");
            worksheet.Cells["A3"].PutValue(2);
            worksheet.Cells["B3"].PutValue("Mary");
            worksheet.Cells["A4"].PutValue(3);
            worksheet.Cells["B4"].PutValue("Bob");

            // Create a table (ListObject) covering A1:B4
            int tableIndex = worksheet.ListObjects.Add("A1", "B4", true);
            ListObject table = worksheet.ListObjects[tableIndex];

            // Destination position (e.g., move to row 10, column C)
            int destRow = 9;      // zero‑based index for row 10
            int destColumn = 2;   // zero‑based index for column C

            // Define the source area using the table's current range
            CellArea sourceArea = new CellArea
            {
                StartRow = table.StartRow,
                StartColumn = table.StartColumn,
                EndRow = table.EndRow,
                EndColumn = table.EndColumn
            };

            // Move the range; structured references inside the table are updated automatically
            worksheet.Cells.MoveRange(sourceArea, destRow, destColumn);

            // Save the workbook
            string outputPath = "MovedTable.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
