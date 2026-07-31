// Title: Aspose.Cells for .NET – Enable Cell‑Based Data Labels (ShowCellRange) on a Chart Series (C#)
// Description: Creates a workbook, adds a column chart, sets a series, turns on data labels, activates ShowCellRange, links labels to cells C2:C3, applies blue font, and saves the file.
// Keywords: Aspose.Cells | C# chart data labels | Series.DataLabels.ShowCellRange | LinkedSource property | custom chart labels from cells | column chart Aspose.Cells | Excel workbook .NET | cell‑based data labels | Aspose.Cells .NET example | chart series styling
// Common Searches: Aspose.Cells ShowCellRange example C# | How to enable cell‑based data labels in Aspose.Cells chart | Series.DataLabels.ShowCellRange true usage | Link chart data labels to a cell range Aspose.Cells | Custom data labels from cells Aspose.Cells .NET | Activate ShowCellRange for chart series
// Developer Intent: Activate cell‑based data labels for a chart series and bind them to a specific cell range.
// Use Cases: Display custom text (e.g., units or descriptions) from adjacent cells as data labels on a column chart. | Generate financial or sales reports where labels show formatted strings instead of raw numeric values. | Apply visual styling such as font color or size to cell‑based labels for clearer presentation.
// AI Prompts: Generate C# code using Aspose.Cells to create a line chart, enable ShowCellRange for its series, and link labels to cells D2:D10. | Explain how ShowCellRange, ShowValue, and LinkedSource interact when configuring chart data labels in Aspose.Cells. | Show how to format cell‑based data labels with bold text, background fill, and custom font size in an Aspose.Cells chart.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using System.Drawing;

namespace AsposeCellsShowCellRangeDemo
{
    // Creates a workbook, adds a column chart, sets a series, turns on data labels, activates ShowCellRange, links labels to cells C2:C3, applies blue font, and saves the file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["B2"].PutValue(100);
            sheet.Cells["B3"].PutValue(200);
            sheet.Cells["C2"].PutValue("100 units");
            sheet.Cells["C3"].PutValue("200 units");

            // Add a column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Define series data and categories
            chart.NSeries.Add("B2:B3", true);
            chart.NSeries.CategoryData = "A2:A3";

            // Access the first series
            Series series = chart.NSeries[0];

            // Enable data labels and activate cell‑based labels
            series.DataLabels.ShowValue = true;          // optional: show the numeric value
            series.DataLabels.ShowCellRange = true;      // activate cell range as data labels
            series.DataLabels.LinkedSource = "C2:C3";    // link to cells containing custom label text
            series.DataLabels.Font.Color = Color.Blue;  // optional styling

            // Save the workbook
            workbook.Save("ShowCellRangeDemo.xlsx");
        }
    }
}
