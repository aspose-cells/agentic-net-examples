// Title: C# – Apply a Top 5 AutoFilter with Aspose.Cells
// Description: Shows how to create a workbook, fill a column with numeric values, set an AutoFilter range, and use Aspose.Cells FilterTop10 to keep only the five highest entries before refreshing and saving the file.
// Keywords: Aspose.Cells AutoFilter | C# FilterTop10 | Top 5 filter Aspose | Excel top rows C# | FilterTop10 example | limit rows to top five | Excel AutoFilter C# | Aspose.Cells top items
// Common Searches: Aspose.Cells top 5 filter C# | FilterTop10 keep highest values | How to show only top five rows in Excel using Aspose | C# AutoFilter top items Aspose.Cells | Apply Top10 filter with count 5 Aspose
// Developer Intent: Generate an Excel file and display only the five highest values in a column by applying an AutoFilter via Aspose.Cells.
// Use Cases: Create a leaderboard that lists the top five scores. | Produce a sales report highlighting the five best‑selling products. | Export a dataset containing only the top five performance metrics. | Build a dashboard that automatically filters to the highest five entries. | Prepare a summary sheet that shows the top five results for quick review.
// AI Prompts: Write C# code using Aspose.Cells to apply a Top 5 AutoFilter on column C of an existing workbook and save it. | Show how to modify FilterTop10 to filter the bottom 10 % of values in a pivot table with Aspose.Cells. | Provide a complete example that adds conditional formatting to rows after applying a Top 5 AutoFilter in Aspose.Cells. | Explain how to combine AutoFilter with sorting to display the top five rows in descending order using Aspose.Cells C#. | Generate a script that reads data from a CSV, loads it into Aspose.Cells, applies a Top 5 filter, and exports to XLSX.

using System;
using Aspose.Cells;

namespace AsposeCellsTop5FilterDemo
{
    // Shows how to create a workbook, fill a column with numeric values, set an AutoFilter range, and use Aspose.Cells FilterTop10 to keep only the five highest entries before refreshing and saving the file.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data (header + 10 numeric values)
            sheet.Cells["A1"].PutValue("Score");
            for (int i = 2; i <= 11; i++)
            {
                // Example values; you can replace with your own data
                sheet.Cells[$"A{i}"].PutValue(100 - (i - 2) * 7);
            }

            // Define the autofilter range (including the header row)
            sheet.AutoFilter.Range = "A1:A11";

            // Apply a Top 5 filter on the first column (field index 0)
            // isTop = true  -> keep top items
            // isPercent = false -> use absolute count, not percent
            // itemCount = 5 -> keep the best five items
            sheet.AutoFilter.FilterTop10(fieldIndex: 0, isTop: true, isPercent: false, itemCount: 5);

            // Refresh the filter to hide rows that do not meet the criteria
            sheet.AutoFilter.Refresh();

            // Save the workbook to a file
            workbook.Save("Top5FilterDemo.xlsx");
        }
    }
}
