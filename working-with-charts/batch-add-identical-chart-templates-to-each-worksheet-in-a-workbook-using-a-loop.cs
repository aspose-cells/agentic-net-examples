// Title: Add an identical column chart to each worksheet in an Aspose.Cells workbook using C#
// AI Prompts: Generate C# code that iterates over all worksheets in a Workbook and inserts a column chart using the same data range on each sheet. | Demonstrate how to reuse a chart template for every worksheet in an Aspose.Cells workbook using a foreach loop. | Write a C# snippet that creates sample data, adds a column chart, and saves the workbook with charts on all sheets.
// Common Searches: C# Aspose.Cells add column chart to each sheet in a workbook | Loop through worksheets and insert identical chart using Aspose.Cells .NET | Apply same chart template to multiple worksheets Aspose.Cells example | Batch create charts on all worksheets with Aspose.Cells C#
// Tags: batch chart insertion across worksheets Aspose.Cells | column chart template application C# | worksheet iteration for chart creation Aspose.Cells | sample data population for Excel chart Aspose.Cells | export workbook with charts Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a workbook, adds several worksheets, populates each with sample category/value data, and within a foreach loop adds an identical column chart titled "Sample Column Chart" using the range B2:B6 on every sheet. The workbook is then saved as WorkbookWithCharts.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (contains a default sheet)
            Workbook workbook = new Workbook();

            // Add additional worksheets for demonstration
            workbook.Worksheets.Add("Sheet1");
            workbook.Worksheets.Add("Sheet2");
            workbook.Worksheets.Add("Sheet3");

            // Loop through each worksheet and add identical chart template
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                Cells cells = sheet.Cells;

                // ----- Populate sample data (Category / Value) -----
                cells["A1"].PutValue("Category");
                cells["B1"].PutValue("Value");

                string[] categories = { "A", "B", "C", "D", "E" };
                double[] values = { 10, 20, 30, 25, 15 };

                for (int i = 0; i < categories.Length; i++)
                {
                    cells[i + 2, 0].PutValue(categories[i]); // Column A (Category)
                    cells[i + 2, 1].PutValue(values[i]);    // Column B (Value)
                }

                // ----- Add a Column chart -----
                // Parameters: ChartType, upper-left row, upper-left column, lower-right row, lower-right column
                int chartIdx = sheet.Charts.Add(ChartType.Column, 5, 0, 25, 10);
                Chart chart = sheet.Charts[chartIdx];

                // Set chart title
                chart.Title.Text = "Sample Column Chart";

                // Add series with values range; category data is optional for this demo
                int seriesIdx = chart.NSeries.Add("B2:B6", true);
                // Optional: name the series
                chart.NSeries[seriesIdx].Name = "Values";
            }

            // Define output file path
            string outputPath = "WorkbookWithCharts.xlsx";

            // Ensure the directory exists before saving
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook with charts
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
