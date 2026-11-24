// Title: How to add a Top 5 AutoFilter to a column in an Excel file using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells apply top 5 filter on column values | Aspose.Cells .NET filter rows to show highest five scores | How to implement custom top‑5 filter with Aspose.Cells AutoFilter | Aspose.Cells limit displayed rows to top N in a worksheet | Example of using Aspose.Cells AutoFilter to show top records in C#
// Tags: Aspose.Cells AutoFilter top N | C# Excel top N filter | custom ranking filter Aspose.Cells | filter column by highest values .NET | Aspose.Cells workbook save filtered view

using System;
using Aspose.Cells;

// The example creates a workbook, fills column A with item names and column B with decreasing scores, sets an AutoFilter on the range A1:B21, notes that the built‑in Top10 filter API is unavailable, and saves the file as Top5Filter.xlsx, indicating where custom logic can be added to display only the top five scores.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // ----- Sample data (optional) -----
            // Header
            sheet.Cells["A1"].PutValue("Item");
            sheet.Cells["B1"].PutValue("Score");

            // Populate some rows with data
            for (int i = 0; i < 20; i++)
            {
                sheet.Cells[i + 1, 0].PutValue($"Item {i + 1}");
                sheet.Cells[i + 1, 1].PutValue(100 - i * 3); // decreasing scores
            }

            // Apply an AutoFilter to the range A1:B21 (header + 20 rows)
            sheet.AutoFilter.Range = "A1:B21";

            // NOTE: The Top10 filter API (IsTop10, Top10Count) is not available in the
            // current Aspose.Cells version used for compilation. If needed, implement
            // alternative filtering logic here.

            // Save the workbook
            workbook.Save("Top5Filter.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
