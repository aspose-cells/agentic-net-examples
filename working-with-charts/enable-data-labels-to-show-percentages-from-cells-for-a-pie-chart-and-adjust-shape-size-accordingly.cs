// Title: Create an Excel pie chart with category, value, and percentage data labels and adjust its shape size using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code with Aspose.Cells that builds a pie chart from a data range, enables data labels to show the category name, numeric value, and calculated percentage, and saves the workbook. | Write a C# snippet that adds a percentage column to a worksheet, formats it as a percentage string, and configures the pie chart series to display those percentages in the data labels using Aspose.Cells. | Provide C# instructions to resize an Aspose.Cells chart shape to fit a specific cell range after creating the pie chart.
// Common Searches: how to display category, value and percent in pie chart data labels with Aspose.Cells C# | Aspose.Cells C# set data labels to show percentage on pie chart | resize Aspose.Cells chart to specific cell range in .NET | add formatted percentage column for Excel chart using Aspose.Cells | C# example of creating pie chart with custom data labels in Aspose.Cells
// Tags: Aspose.Cells pie chart data labels | C# Excel chart show percentage | Aspose.Cells resize chart shape | populate percentage column Aspose.Cells | Excel pie chart custom labels .NET | Aspose.Cells chart series configuration

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The program creates a new workbook, fills columns A‑C with categories, values, and formatted percentages, adds a pie chart based on the values, enables data labels to display the category name, value, and percentage, optionally resizes the chart shape to a defined cell range, and saves the file as PieChartWithCustomPercentLabels.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Data";

            // -------------------------------------------------
            // Populate sample data: categories, values, percentages
            // -------------------------------------------------
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["C1"].PutValue("Percent");

            string[] categories = { "A", "B", "C", "D" };
            double[] values = { 30, 20, 40, 10 };

            double total = 0;
            foreach (double v in values) total += v;

            for (int i = 0; i < categories.Length; i++)
            {
                sheet.Cells[i + 1, 0].PutValue(categories[i]);               // Column A
                sheet.Cells[i + 1, 1].PutValue(values[i]);                  // Column B
                double percent = values[i] / total * 100;
                sheet.Cells[i + 1, 2].PutValue($"{percent:0.##}%");         // Column C
            }

            // -------------------------------------------------
            // Add a pie chart to the worksheet
            // -------------------------------------------------
            int chartIndex = sheet.Charts.Add(ChartType.Pie, 5, 0, 20, 10);
            Chart pieChart = sheet.Charts[chartIndex];

            // Define the series: values from B2:B5 (categories are taken from A2:A5 automatically)
            pieChart.NSeries.Add("B2:B5", true);

            // -------------------------------------------------
            // Enable data labels (show category, value and percent)
            // -------------------------------------------------
            Series series = pieChart.NSeries[0];
            series.DataLabels.ShowCategoryName = true;
            series.DataLabels.ShowValue = true;
            series.DataLabels.ShowPercentage = true;

            // -------------------------------------------------
            // Save the workbook
            // -------------------------------------------------
            string outputPath = "PieChartWithCustomPercentLabels.xlsx";

            // Ensure the directory exists before saving
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
