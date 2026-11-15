// Title: Create a combo chart with column and line series in an Excel workbook using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# create combo chart with column and line series | how to add a line series to an existing column chart in Aspose.Cells | C# generate Excel chart with mixed column and line series using Aspose.Cells | set series type to line programmatically with Aspose.Cells | create combo chart from separate data ranges for sales and profit in Aspose.Cells
// Tags: Aspose.Cells create combo chart | add column series Aspose.Cells | convert series to line Aspose.Cells | Excel chart mixed column line C# | set NSeries CategoryData Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;
using System.IO;

// The example builds a new workbook, fills month labels and two data series (sales and profit), adds a column chart, inserts a column series for sales and a line series for profit, changes the second series type to Line to form a combo chart, assigns the category axis data, and saves the file as ComboChart.xlsx.
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

            // Populate categories
            sheet.Cells["A2"].PutValue("Jan");
            sheet.Cells["A3"].PutValue("Feb");
            sheet.Cells["A4"].PutValue("Mar");
            sheet.Cells["A5"].PutValue("Apr");

            // Populate data for the column series
            sheet.Cells["B2"].PutValue(120);
            sheet.Cells["B3"].PutValue(150);
            sheet.Cells["B4"].PutValue(180);
            sheet.Cells["B5"].PutValue(200);

            // Populate data for the line series
            sheet.Cells["C2"].PutValue(30);
            sheet.Cells["C3"].PutValue(45);
            sheet.Cells["C4"].PutValue(60);
            sheet.Cells["C5"].PutValue(80);

            // Add a chart (initially a column chart) to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 7, 0, 25, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Set chart title
            chart.Title.Text = "Combo Chart Example";

            // Add the column series (first series)
            int columnSeriesIdx = chart.NSeries.Add("B2:B5", true);
            chart.NSeries[columnSeriesIdx].Name = "Sales";

            // Add the line series (second series)
            int lineSeriesIdx = chart.NSeries.Add("C2:C5", true);
            chart.NSeries[lineSeriesIdx].Name = "Profit";

            // Change the second series type to Line to create a combo chart
            chart.NSeries[lineSeriesIdx].Type = ChartType.Line;

            // Set the category (X‑axis) data
            chart.NSeries.CategoryData = "A2:A5";

            // Define output file path
            string outputPath = "ComboChart.xlsx";

            // Ensure the directory exists before saving
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
