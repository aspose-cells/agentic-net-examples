// Title: Create multiple Excel tables in one worksheet, assign a TableCssId‑prefixed TableStyleName, and export to XLSX using Aspose.Cells for .NET
// AI Prompts: Generate a .NET workbook, add ListObjects named Table1 and Table2, set each TableStyleName to "TableCssId-<name>", and save the file as XLSX. | Build a worksheet with several data ranges, convert each range into a ListObject, apply a custom style name prefixed with TableCssId, and export the workbook using Aspose.Cells. | Programmatically populate a sheet with multiple tables, assign a CSS‑like identifier to the TableStyleName of each table, and write the workbook to disk in C#.
// Common Searches: Aspose.Cells how to set TableStyleName with a custom prefix for each ListObject | C# export a worksheet that contains multiple tables to a single XLSX file using Aspose.Cells | assign CSS‑like class names to Excel tables created by Aspose.Cells .NET
// Tags: Aspose.Cells ListObject TableStyleName prefix | export multiple tables to XLSX Aspose.Cells .NET | C# create Excel tables with custom style identifier | assign TableCssId style to Excel ListObject | single worksheet multiple ListObjects Aspose.Cells

using System;
using System.Collections.Generic;
using Aspose.Cells;
using Aspose.Cells.Tables;

// Demonstrates creating a workbook, inserting two tables with data, setting each table's TableStyleName to a value prefixed by "TableCssId-", and saving the result as MultipleTables.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet and give it a name
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Data";

            // Prepare data for multiple tables
            var tablesData = new List<(string name, string[,] data)>
            {
                ("Table1", new string[,] { {"ID","Name"},{"1","Alice"},{"2","Bob"} }),
                ("Table2", new string[,] { {"Product","Price"},{"Apple","1.2"},{"Banana","0.8"} })
            };

            int startRow = 0; // Starting row for the first table

            foreach (var (name, data) in tablesData)
            {
                // Fill the data into the worksheet cells
                int rows = data.GetLength(0);
                int cols = data.GetLength(1);
                for (int r = 0; r < rows; r++)
                {
                    for (int c = 0; c < cols; c++)
                    {
                        sheet.Cells[startRow + r, c].PutValue(data[r, c]);
                    }
                }

                // Define the range that will become the table
                CellArea tableArea = new CellArea
                {
                    StartRow = startRow,
                    StartColumn = 0,
                    EndRow = startRow + rows - 1,
                    EndColumn = cols - 1
                };

                // Convert CellArea to an address string (e.g., "A1:C3")
                string startCell = CellsHelper.CellIndexToName(tableArea.StartRow, tableArea.StartColumn);
                string endCell = CellsHelper.CellIndexToName(tableArea.EndRow, tableArea.EndColumn);
                string areaRef = $"{startCell}:{endCell}";

                // Add a ListObject (table) to the worksheet with a name
                int tableIndex = sheet.ListObjects.Add(name, areaRef, true);
                ListObject table = sheet.ListObjects[tableIndex];

                // Set the display name of the table
                table.DisplayName = name;

                // Assign a style name prefixed with "TableCssId-" (as a substitute for CSS ID)
                table.TableStyleName = "TableCssId-" + name;

                // Leave a blank row before the next table
                startRow += rows + 2;
            }

            // Save the workbook to a file
            workbook.Save("MultipleTables.xlsx", SaveFormat.Xlsx);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
