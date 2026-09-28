// Title: How to hide the legend of a 3‑D pie chart in an Excel workbook using Aspose.Cells for .NET
// AI Prompts: Create a new workbook, populate category and value cells, add a 3‑D pie chart, set ShowLegend = false, give the chart a title, and save the file with Aspose.Cells in C#. | Modify an existing 3‑D pie chart to remove its legend and assign a custom title by using the Aspose.Cells chart API in C#.
// Common Searches: Aspose.Cells C# hide legend on 3D pie chart example | remove legend from 3D pie chart generated with Aspose.Cells .NET | how to disable chart legend for a 3D pie chart using Aspose.Cells | effect of hiding legend on readability of 3D pie charts in Excel with Aspose.Cells
// Tags: Aspose.Cells hide chart legend | 3D pie chart ShowLegend property | Aspose.Cells chart title configuration | Excel workbook create 3D pie chart C# | Aspose.Cells chart readability optimization

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;

// The sample creates a workbook, writes category and value data, inserts a 3‑D pie chart, disables its legend via the ShowLegend property, sets a chart title, and saves the workbook as an XLSX file using Aspose.Cells for .NET.
class HideLegend3DPieChart
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate data for the pie chart (Categories in column A, Values in column B)
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("Apple");
            sheet.Cells["B2"].PutValue(30);
            sheet.Cells["A3"].PutValue("Banana");
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["A4"].PutValue("Cherry");
            sheet.Cells["B4"].PutValue(25);
            sheet.Cells["A5"].PutValue("Date");
            sheet.Cells["B5"].PutValue(25);

            // Add a 3‑D pie chart (positioned at row 7, column 0, spanning 15 rows and 10 columns)
            int chartIndex = sheet.Charts.Add(ChartType.Pie3D, 6, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data source for the chart (values and categories)
            chart.NSeries.Add("B2:B5", true);
            chart.NSeries.CategoryData = "A2:A5";

            // Hide the legend (use Chart.ShowLegend property)
            chart.ShowLegend = false;

            // Set a title for the chart
            chart.Title.Text = "Fruit Distribution (3‑D Pie)";

            // Save the workbook
            workbook.Save("HideLegend3DPieChart.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
