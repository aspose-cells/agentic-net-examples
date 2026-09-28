// Title: How to link data label number format to source cells for the first series in an Aspose.Cells column chart (C#)
// AI Prompts: Write C# code that creates a workbook, adds a column chart, enables data labels for the first series, sets DataLabels.NumberFormatLinked = true, and saves the file. | Show an example of binding a chart series' data label number format to a range of formatted cells using Aspose.Cells for .NET.
// Common Searches: Aspose.Cells set DataLabels.NumberFormatLinked for first series in C# | link chart data label formatting to worksheet cells Aspose.Cells .NET | C# Aspose.Cells column chart data labels use source cell number format | how to bind data label number format to cells in Aspose.Cells chart
// Tags: Aspose.Cells chart data label formatting binding | C# enable number format linking for chart series | Aspose.Cells column chart series data label settings | link data label number format to source range Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsDemo
{
    // Demonstrates creating a workbook, adding a column chart, populating sample data, enabling data labels, linking them to formatted cells, setting DataLabels.NumberFormatLinked = true for the first series, and saving the workbook as an XLSX file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["B2"].PutValue(100);
            sheet.Cells["B3"].PutValue(200);
            // Formatted values in another column (optional for linking)
            sheet.Cells["C1"].PutValue("Formatted Value");
            sheet.Cells["C2"].PutValue("100 units");
            sheet.Cells["C3"].PutValue("200 units");

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Define the data range for the series and categories
            chart.NSeries.Add("B2:B3", true);          // Values
            chart.NSeries.CategoryData = "A2:A3";      // Categories

            // Access the first series
            Series firstSeries = chart.NSeries[0];

            // Enable data labels and link them to the formatted cells (optional)
            firstSeries.DataLabels.ShowValue = true;
            firstSeries.DataLabels.LinkedSource = "C2:C3";

            // Bind the number format of the data labels to the source cells
            firstSeries.DataLabels.NumberFormatLinked = true;

            // Save the workbook to an XLSX file
            workbook.Save("DataLabelsNumberFormatLinkedDemo.xlsx");
        }
    }
}
