// Title: Move an Excel Table (ListObject) within a same worksheet using Aspose.Cells for .NET
// Description: Demonstrates how to create a workbook, define a ListObject named "MyTable" over A1:B4, and relocate the entire table to cell D5 with Worksheet.Cells.MoveRange. All formulas, formatting, and structured references are refreshed automatically, and the result is saved as MovedTable.xlsx.
// Keywords: Aspose.Cells move table | C# ListObject relocation | Worksheet Cells MoveRange | preserve structured references | update Excel formulas after move | .NET Excel table reposition
// Common Searches: how to move an Excel table with Aspose.Cells | preserve structured references when moving ListObject .NET | move range with formulas C# Aspose.Cells | relocate Excel ListObject without breaking formulas | Aspose.Cells MoveRange example
// Developer Intent: Reposition an existing ListObject to a new cell range in the same worksheet while letting Aspose.Cells automatically adjust all references.
// Use Cases: Rearrange data tables to free space for new report sections. | Shift a table after inserting rows or columns to keep layout consistent. | Prepare a workbook for export by placing tables in a predefined order.
// AI Prompts: Generate C# code with Aspose.Cells that moves a ListObject from A1:B4 to D5 and updates structured references. | Explain how Worksheet.Cells.MoveRange rewrites formulas and table references after a table is moved. | Show best‑practice error handling for moving tables using Aspose.Cells in a .NET application.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;
using AsposeRange = Aspose.Cells.Range; // Alias to avoid conflict with System.Range

// Demonstrates how to create a workbook, define a ListObject named "MyTable" over A1:B4, and relocate the entire table to cell D5 with Worksheet.Cells.MoveRange. All formulas, formatting, and structured references are refreshed automatically, and the result is saved as MovedTable.xlsx.
class MoveTableDemo
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data that will become a table
            worksheet.Cells["A1"].PutValue("ID");
            worksheet.Cells["B1"].PutValue("Name");
            worksheet.Cells["A2"].PutValue(1);
            worksheet.Cells["B2"].PutValue("John");
            worksheet.Cells["A3"].PutValue(2);
            worksheet.Cells["B3"].PutValue("Mary");
            worksheet.Cells["A4"].PutValue(3);
            worksheet.Cells["B4"].PutValue("Bob");

            // Create a ListObject (Excel table) covering the range A1:B4
            int tableIndex = worksheet.ListObjects.Add("A1", "B4", true);
            ListObject table = worksheet.ListObjects[tableIndex];

            // Set a display name for the table
            table.DisplayName = "MyTable";

            // Define the source area of the table using its data range
            AsposeRange dataRange = table.DataRange;
            CellArea sourceArea = new CellArea
            {
                StartRow = dataRange.FirstRow,
                StartColumn = dataRange.FirstColumn,
                EndRow = dataRange.FirstRow + dataRange.RowCount - 1,
                EndColumn = dataRange.FirstColumn + dataRange.ColumnCount - 1
            };

            // Destination top‑left cell (zero‑based indices). D5 => row 4, column 3
            int destRow = 4;      // Row index for row 5
            int destColumn = 3;   // Column index for column D

            // Move the range; formulas, formatting, and structured references are updated automatically
            worksheet.Cells.MoveRange(sourceArea, destRow, destColumn);

            // Save the workbook
            string outputPath = "MovedTable.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
