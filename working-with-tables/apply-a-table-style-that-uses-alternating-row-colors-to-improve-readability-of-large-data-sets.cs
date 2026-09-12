// Title: Apply a built-in banded-row table style to a 100-row Excel worksheet using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code with Aspose.Cells that creates a 100-row, 5-column worksheet, adds a ListObject, and applies TableStyleMedium2 with row stripes. | Show how to enable row stripe visibility on an Aspose.Cells ListObject while disabling column stripes. | Provide a complete example that saves the styled table to an .xlsx file and outputs a success message.
// Common Searches: Aspose.Cells C# how to add a table with alternating row colors | C# Aspose.Cells apply built-in table style with banded rows | Set ShowTableStyleRowStripes property in Aspose.Cells ListObject example | Create Excel table with TableStyleMedium2 using Aspose.Cells .NET | Save styled Excel table to file with Aspose.Cells C#
// Tags: Aspose.Cells apply TableStyleMedium2 | C# ListObject alternating row colors | Aspose.Cells enable row stripes | Excel table styling with Aspose.Cells .NET | save styled worksheet Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Tables;   // Required for ListObject and TableStyleType

// The example creates a new workbook, fills a 100‑row by 5‑column range with sample data, adds a ListObject covering that range, sets TableStyleType to TableStyleMedium2, enables row stripe display, disables column stripes, and saves the workbook as StyledTable.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook.
            Workbook workbook = new Workbook();

            // Access the first worksheet.
            Worksheet sheet = workbook.Worksheets[0];

            // Populate the worksheet with sample data (e.g., 100 rows, 5 columns).
            int totalRows = 100;
            int totalColumns = 5;
            for (int row = 0; row < totalRows; row++)
            {
                for (int col = 0; col < totalColumns; col++)
                {
                    // Fill each cell with a simple value for demonstration.
                    sheet.Cells[row, col].PutValue($"R{row + 1}C{col + 1}");
                }
            }

            // Define the range that will become a table (list object).
            int firstRow = 0;
            int firstColumn = 0;
            int lastRow = totalRows - 1;
            int lastColumn = totalColumns - 1;

            // Add a ListObject (table) to the worksheet.
            int tableIndex = sheet.ListObjects.Add(firstRow, firstColumn, lastRow, lastColumn, true);
            ListObject table = sheet.ListObjects[tableIndex];

            // Apply a built‑in table style that includes alternating row colors (banded rows).
            table.TableStyleType = TableStyleType.TableStyleMedium2;

            // Ensure that row stripes (alternating colors) are displayed.
            table.ShowTableStyleFirstColumn = false;
            table.ShowTableStyleLastColumn = false;
            table.ShowTableStyleRowStripes = true;
            table.ShowTableStyleColumnStripes = false;

            // Save the workbook to a file.
            string outputPath = "StyledTable.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
