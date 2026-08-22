// Title: How to apply a Top 5 AutoFilter to a column in an Excel worksheet using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that creates a workbook, adds sample data, sets an AutoFilter on a range, and uses FilterTop10 to keep only the five highest values. | Show the steps to programmatically apply a Top 5 filter on the first column of an Excel sheet with Aspose.Cells, including refreshing and saving the file. | Generate a complete Aspose.Cells example that demonstrates populating a column, configuring an AutoFilter, applying a top‑five filter, and exporting the result to a .xlsx file.
// Common Searches: Aspose.Cells C# how to filter top 5 rows in a column | C# using FilterTop10 method to display only highest five values in Excel | Apply AutoFilter with top 5 criteria in Aspose.Cells .NET example | Programmatically limit Excel column to top five entries using Aspose.Cells | Save workbook after applying top five AutoFilter with Aspose.Cells C#
// Tags: Aspose.Cells AutoFilter top‑five | C# FilterTop10 usage | Excel top‑values filter with Aspose.Cells | AutoFilter on .xlsx via C# | limit rows displayed using Aspose.Cells

using System;
using Aspose.Cells;

namespace AsposeCellsTop5FilterDemo
{
    // The sample creates a new workbook, fills column A with a header and ten numeric scores, defines an AutoFilter over the range A1:A11, applies a Top 5 filter on the first column using the FilterTop10 method, refreshes the filter to hide non‑matching rows, and saves the result as Top5FilterDemo.xlsx.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data in column A (header + 10 numeric values)
            sheet.Cells["A1"].PutValue("Score");
            for (int i = 2; i <= 11; i++)
            {
                // Example values; you can replace with your own data
                sheet.Cells[$"A{i}"].PutValue(100 - (i - 2) * 7);
            }

            // Apply an AutoFilter to the data range (including header)
            sheet.AutoFilter.Range = "A1:A11";

            // Apply a Top 5 filter on the first column (field index 0)
            // Parameters: fieldIndex, isTop, isPercent, itemCount
            sheet.AutoFilter.FilterTop10(fieldIndex: 0, isTop: true, isPercent: false, itemCount: 5);

            // Refresh the filter to hide rows that do not meet the criteria
            sheet.AutoFilter.Refresh();

            // Save the workbook
            workbook.Save("Top5FilterDemo.xlsx");
        }
    }
}
