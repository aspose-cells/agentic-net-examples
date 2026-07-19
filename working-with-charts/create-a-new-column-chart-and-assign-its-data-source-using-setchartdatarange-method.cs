// Title: Aspose.Cells for .NET: Create a Column Chart and Bind Data with SetChartDataRange (C#)
// Description: Learn how to generate a column chart in a new workbook, fill cells A1:B4 with categories and values, and link the chart to that range using the SetChartDataRange method (plot by column). The example saves the result as an XLSX file.
// Keywords: Aspose.Cells C# SetChartDataRange | column chart Aspose.Cells .NET | chart data source programmatically | Excel chart API C# | Aspose.Cells chart example
// Common Searches: Aspose.Cells SetChartDataRange column chart example | C# create column chart from worksheet data | bind Excel chart to range using Aspose.Cells | how to plot chart by column in Aspose.Cells | set chart data source programmatically C#
// Developer Intent: Create a column chart in a workbook and attach its category and value series by specifying a cell range with SetChartDataRange.
// Use Cases: Generate a sales‑by‑region column chart for monthly reporting. | Build an automated dashboard where chart data updates with worksheet changes. | Export a formatted column chart for inclusion in presentations or client deliverables.
// AI Prompts: Provide C# code that creates a stacked column chart and sets its data range using SetChartDataRange in Aspose.Cells. | Explain how to modify SetChartDataRange to plot by rows instead of columns for a column chart. | Show how to add multiple series to a column chart by calling SetChartDataRange with separate ranges.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartDemo
{
    // Learn how to generate a column chart in a new workbook, fill cells A1:B4 with categories and values, and link the chart to that range using the SetChartDataRange method (plot by column). The example saves the result as an XLSX file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Series1");
            sheet.Cells["A2"].PutValue("Cat1");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["A3"].PutValue("Cat2");
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["A4"].PutValue("Cat3");
            sheet.Cells["B4"].PutValue(30);

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
            Chart chart = sheet.Charts[chartIndex];

            // Assign the data source to the chart using SetChartDataRange
            // The range includes both category (A column) and values (B column)
            // The second argument 'true' indicates plotting by column
            chart.SetChartDataRange("A1:B4", true);

            // Optional: set a title for clarity
            chart.Title.Text = "Sample Column Chart";

            // Save the workbook to an XLSX file
            workbook.Save("ColumnChartWithDataRange.xlsx", SaveFormat.Xlsx);
        }
    }
}
