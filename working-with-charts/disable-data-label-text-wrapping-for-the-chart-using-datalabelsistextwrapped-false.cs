// Title: Aspose.Cells C# – Turn Off Chart Data Label Text Wrapping (DataLabels.IsTextWrapped = false)
// Description: Creates a workbook, adds a column chart with sales data, shows values on the first series' data labels, and disables label text wrapping by setting DataLabels.IsTextWrapped to false before saving as ChartDataLabels_NoWrap.xlsx.
// Keywords: Aspose.Cells chart data labels | DataLabels.IsTextWrapped | C# disable label wrap | Excel chart label formatting .NET | Aspose.Cells column chart example | prevent text wrap in chart labels | Aspose.Cells API DataLabels
// Common Searches: Aspose.Cells set DataLabels.IsTextWrapped false | C# chart label no wrap Aspose | how to stop text wrapping on Excel chart labels using Aspose.Cells | disable data label wrap in column chart .NET | Aspose.Cells chart label formatting options
// Developer Intent: Disable automatic line‑breaks in chart data label text.
// Use Cases: Generate reports where chart labels must stay on a single line for readability. | Create Excel files with long numeric or textual values in labels without disturbing layout. | Prepare spreadsheets for presentation where wrapped labels would cause misalignment.
// AI Prompts: Provide C# code using Aspose.Cells to add a bar chart and set DataLabels.IsTextWrapped = false for all series. | Explain how DataLabels.IsTextWrapped interacts with ShowValue and other label properties in Aspose.Cells charts. | Show how to programmatically ensure chart data labels never wrap, regardless of label length, in a .NET workbook.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartDataLabelWrapDemo
{
    // Creates a workbook, adds a column chart with sales data, shows values on the first series' data labels, and disables label text wrapping by setting DataLabels.IsTextWrapped to false before saving as ChartDataLabels_NoWrap.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            worksheet.Cells["A1"].PutValue("Category");
            worksheet.Cells["A2"].PutValue("Apple");
            worksheet.Cells["A3"].PutValue("Orange");
            worksheet.Cells["A4"].PutValue("Banana");
            worksheet.Cells["B1"].PutValue("Sales");
            worksheet.Cells["B2"].PutValue(120);
            worksheet.Cells["B3"].PutValue(85);
            worksheet.Cells["B4"].PutValue(65);

            // Add a column chart to the worksheet
            int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart chart = worksheet.Charts[chartIndex];

            // Set the data range for the series and categories
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Access the data labels of the first series
            DataLabels dataLabels = chart.NSeries[0].DataLabels;

            // Show the values on the data labels
            dataLabels.ShowValue = true;

            // Disable text wrapping for the data labels
            dataLabels.IsTextWrapped = false;

            // Save the workbook to a file
            workbook.Save("ChartDataLabels_NoWrap.xlsx");
        }
    }
}
