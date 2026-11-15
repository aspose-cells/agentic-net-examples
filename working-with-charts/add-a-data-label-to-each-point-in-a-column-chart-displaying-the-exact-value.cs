// Title: How to add data labels that show exact values to a clustered column chart using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# add data labels with values to a column chart | show exact series values on each column in an Excel chart using Aspose.Cells | C# code to enable data labels for a clustered column chart in Aspose.Cells | Aspose.Cells .NET set DataLabels.ShowValue for column chart series | create Excel column chart with value labels using Aspose.Cells library | display numeric values on chart points in Aspose.Cells C#
// Tags: Aspose.Cells column chart data labels | DataLabels.ShowValue property C# | clustered column chart Aspose.Cells | Excel chart value labels Aspose.Cells | add series values to chart points .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExample
{
    // The example creates a new workbook, fills cells with month and sales data, adds a clustered column chart referencing the sales range, enables data labels to display each column's exact value via the DataLabels.ShowValue property, and saves the workbook as ColumnChartWithDataLabels.xlsx.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook.
                Workbook workbook = new Workbook();

                // Access the first worksheet.
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for the column chart.
                sheet.Cells["A1"].PutValue("Month");
                sheet.Cells["B1"].PutValue("Sales");
                sheet.Cells["A2"].PutValue("Jan");
                sheet.Cells["A3"].PutValue("Feb");
                sheet.Cells["A4"].PutValue("Mar");
                sheet.Cells["B2"].PutValue(120);
                sheet.Cells["B3"].PutValue(150);
                sheet.Cells["B4"].PutValue(130);

                // Add a clustered column chart to the worksheet.
                int chartIndex = sheet.Charts.Add(ChartType.Column, 6, 0, 20, 7);
                Chart chart = sheet.Charts[chartIndex];

                // Set the data range for the series.
                chart.NSeries.Add("B2:B4", true);

                // Enable data labels to display the exact value for each point.
                chart.NSeries[0].DataLabels.ShowValue = true;

                // Determine output path and ensure the directory exists.
                string outputPath = "ColumnChartWithDataLabels.xlsx";
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook with the chart.
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
