// Title: Create a ListObject table with a frozen header row in an Excel worksheet using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# add ListObject table and freeze first row | How to freeze pane at A2 after creating a structured table with Aspose.Cells | Apply built‑in TableStyleMedium9 to a ListObject using Aspose.Cells .NET | Generate sample data and create a table programmatically with Aspose.Cells | Save workbook with a frozen header row using Aspose.Cells for .NET
// Tags: create ListObject table Aspose.Cells | freeze header row Aspose.Cells | apply TableStyleMedium9 Aspose.Cells | populate worksheet sample data C# | save workbook as XLSX Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AsposeCellsTableFreezeDemo
{
    // The example creates a new workbook, adds a worksheet named "DataSheet", fills cells A1:C11 with header and sample rows, inserts a ListObject named "StudentsTable" covering that range, applies the built‑in TableStyleMedium9, freezes the pane at A2 so the header stays visible while scrolling, and saves the file as TableWithFrozenHeader.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Name = "DataSheet";

                // Populate sample data (including header row)
                sheet.Cells["A1"].PutValue("ID");
                sheet.Cells["B1"].PutValue("Name");
                sheet.Cells["C1"].PutValue("Score");

                // Data rows (rows 2..11)
                for (int i = 2; i <= 11; i++)
                {
                    sheet.Cells[$"A{i}"].PutValue(i - 1); // ID
                    sheet.Cells[$"B{i}"].PutValue($"Student {i - 1}");
                    sheet.Cells[$"C{i}"].PutValue(50 + (i - 2) * 5); // Sample score
                }

                // Define the range that will become a structured table (A1:C11)
                int firstRow = 0;       // zero‑based index for row 1
                int firstColumn = 0;    // column A
                int totalRows = 11;     // rows 1‑11 inclusive
                int totalColumns = 3;   // columns A‑C

                // Add the ListObject (structured table) to the worksheet
                int tableIndex = sheet.ListObjects.Add(firstRow, firstColumn, totalRows - 1, totalColumns - 1, true);
                ListObject table = sheet.ListObjects[tableIndex];

                // Set table properties (use correct property names for the API version)
                table.DisplayName = "StudentsTable";
                table.ShowHeaderRow = true;
                table.ShowTableStyleFirstColumn = false;
                table.ShowTableStyleLastColumn = false;
                table.ShowTableStyleRowStripes = true;
                table.ShowTableStyleColumnStripes = false;
                table.TableStyleType = TableStyleType.TableStyleMedium9; // built‑in style

                // Freeze the header row so it stays visible while scrolling
                sheet.FreezePanes("A2", 1, 0);

                // Save the workbook to a file
                string outputPath = "TableWithFrozenHeader.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
                Console.WriteLine("The table 'StudentsTable' has been created and the header row is frozen.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred while processing the workbook:");
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}
