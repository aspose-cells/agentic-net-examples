// Title: Aspose.Cells for .NET – Insert a Column Chart into a 3‑row × 5‑column range
// Description: Creates a workbook, adds sample data, and places a column chart that starts at cell B3 and spans exactly three rows and five columns. The chart’s series and categories are set, a title is added, and the file is saved as ChartInRange.xlsx.
// Keywords: Aspose.Cells chart placement | C# insert chart by cell range | set chart size rows columns Aspose | column chart specific area | chart dimensions Aspose.Cells
// Common Searches: Aspose.Cells place chart in defined cell range C# | how to size chart 3 rows 5 columns Aspose | chart positioning using topRow leftColumn Aspose.Cells | C# Aspose.Cells chart area coordinates
// Developer Intent: Add a column chart that occupies a predefined block of cells (3 rows × 5 columns) starting from a chosen top‑left cell.
// Use Cases: Design a sales dashboard where the chart aligns with a fixed table layout. | Automate report generation that requires charts to fit into reserved template zones. | Create reusable workbook templates with a chart slot sized to three rows by five columns.
// AI Prompts: Generate C# code with Aspose.Cells to insert a line chart that fits inside cells D5:G7 (3 × 4). | Explain the calculation of bottomRow and rightColumn for a chart covering N rows and M columns. | Provide robust error handling for chart placement when the target range exceeds worksheet limits.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartExample
{
    // Creates a workbook, adds sample data, and places a column chart that starts at cell B3 and spans exactly three rows and five columns. The chart’s series and categories are set, a title is added, and the file is saved as ChartInRange.xlsx.
    public class InsertChartInRange
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for the chart
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Value");
                for (int i = 2; i <= 5; i++)
                {
                    sheet.Cells[$"A{i}"].PutValue($"Item {i - 1}");
                    sheet.Cells[$"B{i}"].PutValue(i * 10);
                }

                // Define the top‑left cell of the chart (row index 2, column index 1 -> cell B3)
                int topRow = 2;      // zero‑based index, corresponds to row 3 in Excel
                int leftColumn = 1;  // zero‑based index, corresponds to column B

                // Chart should occupy exactly 3 rows and 5 columns
                int bottomRow = topRow + 2;      // 3 rows total
                int rightColumn = leftColumn + 4; // 5 columns total

                // Add a column chart positioned in the specified range
                int chartIndex = sheet.Charts.Add(ChartType.Column, topRow, leftColumn, bottomRow, rightColumn);
                Chart chart = sheet.Charts[chartIndex];

                // Set the data source for the chart
                chart.NSeries.Add("B2:B5", true);          // Values
                chart.NSeries.CategoryData = "A2:A5";      // Categories

                // Optional: give the chart a title
                chart.Title.Text = "Sample Column Chart";

                // Save the workbook
                workbook.Save("ChartInRange.xlsx", SaveFormat.Xlsx);
                Console.WriteLine("Workbook saved as ChartInRange.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            InsertChartInRange.Run();
        }
    }
}
