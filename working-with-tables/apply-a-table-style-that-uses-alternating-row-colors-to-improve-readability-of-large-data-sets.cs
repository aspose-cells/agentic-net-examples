// Title: Apply alternating row colors to an Excel table using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells how to add banded rows to an Excel table | set alternating row colors in a ListObject with Aspose.Cells .NET | use TableStyleMedium2 to enable row stripes in Aspose.Cells workbook | Aspose.Cells example for creating a table with row stripe formatting in C#
// Tags: Aspose.Cells apply table style banded rows | C# ListObject ShowTableStyleRowStripes | TableStyleMedium2 Excel table formatting Aspose | create Excel table with alternating row colors Aspose.Cells | Aspose.Cells .NET row stripe table style

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;

// The sample creates a new workbook, fills it with data, adds a ListObject covering A1:D10, sets the table's style to TableStyleMedium2, enables row stripe banding via ShowTableStyleRowStripes, and saves the file as AlternatingRows.xlsx.
class ApplyAlternatingRowColors
{
    static void Main()
    {
        try
        {
            // Create a new workbook.
            Workbook workbook = new Workbook();

            // Access the first worksheet.
            Worksheet sheet = workbook.Worksheets[0];

            // Populate the worksheet with sample data (10 rows, 4 columns).
            for (int row = 0; row < 10; row++)
            {
                if (row == 0) // Header row.
                {
                    sheet.Cells[row, 0].PutValue("ID");
                    sheet.Cells[row, 1].PutValue("Name");
                    sheet.Cells[row, 2].PutValue("Category");
                    sheet.Cells[row, 3].PutValue("Amount");
                }
                else
                {
                    sheet.Cells[row, 0].PutValue(row);                     // ID
                    sheet.Cells[row, 1].PutValue($"Item {row}");           // Name
                    sheet.Cells[row, 2].PutValue($"Group {row % 3 + 1}"); // Category
                    sheet.Cells[row, 3].PutValue(row * 10.5);              // Amount
                }
            }

            // Define the range that will become a table (including headers).
            int firstRow = 0;   // A1
            int firstColumn = 0;
            int lastRow = 9;    // Row 10 (0‑based index)
            int lastColumn = 3; // Column D

            // Add a ListObject (table) to the worksheet.
            int tableIndex = sheet.ListObjects.Add(firstRow, firstColumn, lastRow, lastColumn, true);
            ListObject table = sheet.ListObjects[tableIndex];

            // Apply a built‑in table style that uses alternating row colors.
            table.TableStyleType = TableStyleType.TableStyleMedium2;

            // Enable banded rows (alternating colors) explicitly.
            table.ShowTableStyleFirstColumn = false;
            table.ShowTableStyleLastColumn = false;
            table.ShowTableStyleRowStripes = true;   // ensures alternating row colors
            table.ShowTableStyleColumnStripes = false;

            // Determine output path and ensure the directory exists.
            string outputPath = "AlternatingRows.xlsx";
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook.
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
