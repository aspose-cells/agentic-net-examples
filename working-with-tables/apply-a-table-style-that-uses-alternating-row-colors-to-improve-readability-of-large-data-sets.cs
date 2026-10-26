// Title: Apply alternating row colors (banded rows) to an Excel table with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells how to enable banded rows in an Excel ListObject | Set ShowTableStyleRowStripes property using Aspose.Cells .NET | Apply TableStyleMedium2 to a worksheet table with Aspose.Cells | Create an Excel table with alternating row colors in C# | Aspose.Cells example for styling Excel tables with row banding
// Tags: Aspose.Cells ListObject row banding | C# apply TableStyleMedium2 | Excel table alternating colors Aspose.Cells | Set ShowTableStyleRowStripes in .NET | Create styled ListObject with Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

// Demonstrates creating a new workbook, populating it with sample data, adding a ListObject (Excel table), assigning a display name, applying the built‑in TableStyleMedium2, enabling row banding via ShowTableStyleRowStripes, and saving the result as AlternatingRowColors.xlsx.
class ApplyAlternatingRowColors
{
    static void Main()
    {
        try
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();

            // Access the first worksheet and give it a friendly name
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Data";

            // Populate the worksheet with sample data (header + 100 rows)
            sheet.Cells["A1"].PutValue("ID");
            sheet.Cells["B1"].PutValue("Name");
            sheet.Cells["C1"].PutValue("Score");

            Random rnd = new Random();
            for (int i = 2; i <= 101; i++) // rows 2‑101 contain data
            {
                sheet.Cells[i, 0].PutValue(i - 1);               // ID
                sheet.Cells[i, 1].PutValue($"Item {i - 1}");    // Name
                sheet.Cells[i, 2].PutValue(rnd.Next(0, 100));   // Score
            }

            // Define the range that will become a table (ListObject)
            int totalRows = 101; // header + 100 data rows
            int totalCols = 3;   // ID, Name, Score
            CellArea tableArea = new CellArea
            {
                StartRow = 0,
                StartColumn = 0,
                EndRow = totalRows - 1,
                EndColumn = totalCols - 1
            };

            // Add a ListObject (table) to the worksheet
            int tableIndex = sheet.ListObjects.Add(
                tableArea.StartRow,
                tableArea.StartColumn,
                tableArea.EndRow,
                tableArea.EndColumn,
                true);

            ListObject table = sheet.ListObjects[tableIndex];

            // Set a display name for the table (Name property is not available in this API version)
            table.DisplayName = "DataTable";

            // Apply a built‑in table style that includes alternating (banded) row colors
            table.TableStyleType = TableStyleType.TableStyleMedium2;

            // Enable row banding (alternating colors)
            table.ShowTableStyleRowStripes = true;

            // Save the workbook (lifecycle rule: save)
            string outputPath = "AlternatingRowColors.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
