// Title: Apply Built‑In Chart Style 20 to a Column Chart with Aspose.Cells for .NET
// Description: Creates a workbook, adds sample data, inserts a column chart, sets the built‑in chart style #20 via Chart.Style, and saves the file as ChartWithStyle20.xlsx using Aspose.Cells for .NET.
// Keywords: Aspose.Cells chart style 20 | apply built‑in chart style .NET | Chart.Style property | column chart formatting Aspose | predefined chart styles 1‑48 | C# Aspose.Cells example
// Common Searches: how to set chart style 20 in Aspose.Cells | Aspose.Cells apply built‑in chart style | C# set predefined chart style for column chart | Aspose.Cells chart style range 1 to 48 | change chart appearance programmatically Aspose
// Developer Intent: Set a predefined built‑in chart style (Style20) on a chart programmatically.
// Use Cases: Standardize the look of column charts across multiple reports by applying Chart.Style = 20. | Create a template workbook where every generated chart automatically uses the same visual theme. | Allow end‑users to switch between predefined styles by mapping UI selections to integer style IDs.
// AI Prompts: Show how to apply built‑in chart style 15 to a line chart using Aspose.Cells for .NET. | Generate C# code that lists all available chart style numbers (1‑48) and picks one based on a condition. | Explain how to override series colors after setting a built‑in chart style in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // Creates a workbook, adds sample data, inserts a column chart, sets the built‑in chart style #20 via Chart.Style, and saves the file as ChartWithStyle20.xlsx using Aspose.Cells for .NET.
    public class ApplyChartStyle20
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook (lifecycle rule: create)
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for the chart
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Value");
                sheet.Cells["A2"].PutValue("A");
                sheet.Cells["B2"].PutValue(10);
                sheet.Cells["A3"].PutValue("B");
                sheet.Cells["B3"].PutValue(20);
                sheet.Cells["A4"].PutValue("C");
                sheet.Cells["B4"].PutValue(30);

                // Add a column chart
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
                Chart chart = sheet.Charts[chartIndex];

                // Set data series and categories
                chart.NSeries.Add("B2:B4", false);
                chart.NSeries.CategoryData = "A2:A4";

                // Apply the predefined built‑in chart style #20 (valid range 1‑48)
                chart.Style = 20;

                // Save the workbook (lifecycle rule: save)
                workbook.Save("ChartWithStyle20.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the console application
    public class Program
    {
        public static void Main(string[] args)
        {
            ApplyChartStyle20.Run();
        }
    }
}
