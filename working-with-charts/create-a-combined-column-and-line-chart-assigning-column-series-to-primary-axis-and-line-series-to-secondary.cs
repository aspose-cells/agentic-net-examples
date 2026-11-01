// Title: Create a combined column and line chart with primary and secondary axes in Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code using Aspose.Cells that builds a combo chart where the column series is plotted on the primary axis and the line series on the secondary axis. | Modify the example to change the line series to a spline type and add data labels to both the column and line series. | Extend the program to export the workbook both as an XLSX file and as a PDF while keeping the combined chart intact.
// Common Searches: aspocells c# combined column and line chart with secondary axis example | how to assign a line series to the secondary axis in an Aspose.Cells chart | save Aspose.Cells workbook containing a combo chart as PDF
// Tags: Aspose.Cells combo column line chart C# | set secondary axis for line series Aspose.Cells | save workbook as PDF with chart Aspose.Cells | format chart axes Aspose.Cells .NET | customize chart series appearance Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace CombinedChartExample
{
    // The example creates a new workbook, fills cells A1:C5 with month names and sample data, adds a column chart, adds a column series (B2:B5) on the primary axis and a line series (C2:C5) on the secondary axis, switches the second series to a line type, sets axis titles, and saves the file as CombinedColumnLineChart.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data
                // Column series data (Primary axis)
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Column Series");
                sheet.Cells["C1"].PutValue("Line Series");

                sheet.Cells["A2"].PutValue("Jan");
                sheet.Cells["A3"].PutValue("Feb");
                sheet.Cells["A4"].PutValue("Mar");
                sheet.Cells["A5"].PutValue("Apr");

                sheet.Cells["B2"].PutValue(120);
                sheet.Cells["B3"].PutValue(150);
                sheet.Cells["B4"].PutValue(170);
                sheet.Cells["B5"].PutValue(130);

                sheet.Cells["C2"].PutValue(30);
                sheet.Cells["C3"].PutValue(45);
                sheet.Cells["C4"].PutValue(40);
                sheet.Cells["C5"].PutValue(55);

                // Add a chart to the worksheet
                int chartIndex = sheet.Charts.Add(ChartType.Column, 7, 0, 25, 10);
                Chart chart = sheet.Charts[chartIndex];

                // Set chart title (optional)
                chart.Title.Text = "Combined Column and Line Chart";

                // Add column series (primary axis)
                // Series range: B2:B5 (values), A2:A5 (categories)
                chart.NSeries.Add("B2:B5", true);
                chart.NSeries[0].Name = "Column Series";

                // Add line series (secondary axis)
                // Series range: C2:C5 (values), A2:A5 (categories)
                chart.NSeries.Add("C2:C5", true);
                chart.NSeries[1].Name = "Line Series";

                // Change the second series to a line chart type
                chart.NSeries[1].Type = ChartType.Line;

                // Optionally format primary axes
                chart.CategoryAxis.Title.Text = "Month";
                chart.ValueAxis.Title.Text = "Column Values";

                // Define output file path
                string outputPath = "CombinedColumnLineChart.xlsx";

                // Save the workbook to a file
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
