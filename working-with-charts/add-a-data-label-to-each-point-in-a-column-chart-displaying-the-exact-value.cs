// Title: Add value data labels outside the end of each column in an Aspose.Cells .NET column chart
// AI Prompts: Generate C# code with Aspose.Cells that creates a column chart and enables data labels to show each column's numeric value. | Update an existing Aspose.Cells column chart in C# to turn on value labels and set their position to OutsideEnd.
// Common Searches: how to display data labels with values on a column chart using Aspose.Cells for .NET | C# Aspose.Cells set chart series data label position to OutsideEnd | add exact numeric labels to Excel column chart programmatically with Aspose.Cells
// Tags: Aspose.Cells column chart data labels | C# set data label value Aspose.Cells | Aspose.Cells label position OutsideEnd | Excel column chart value labels .NET | programmatic chart labeling Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // The example creates a new workbook, fills cells A1:B4 with categories and values, adds a column chart, binds the series to the data range, enables data labels to show each point's value, positions the labels outside the end of the columns, and saves the file as ColumnChartWithDataLabels.xlsx.
    public class ColumnChartDataLabels
    {
        public static void Main(string[] args)
        {
            try
            {
                Run();
                Console.WriteLine("Workbook created successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the column chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["A4"].PutValue("C");

            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the series and categories
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Enable data labels for the first series and show the exact value for each point
            chart.NSeries[0].DataLabels.ShowValue = true;
            // Position the labels outside the end of each column
            chart.NSeries[0].DataLabels.Position = LabelPositionType.OutsideEnd;

            // Save the workbook with the chart
            string outputPath = "ColumnChartWithDataLabels.xlsx";
            workbook.Save(outputPath);
        }
    }
}
