// Title: Aspose.Cells for .NET: Bind Chart Data Label Formatting to Source Cells with DataLabels.NumberFormatLinked (C#)
// Description: Demonstrates how to create a workbook, add a column chart, link data labels to formatted cells, and set DataLabels.NumberFormatLinked = true so the label appearance follows the source cell formatting. The example saves the result as an XLSX file.
// Keywords: Aspose.Cells | C# chart example | DataLabels.NumberFormatLinked | link data label format to cells | chart data label formatting | Aspose.Cells .NET tutorial | Excel chart series formatting
// Common Searches: Aspose.Cells set DataLabels.NumberFormatLinked true | link chart data label to cell formatting C# | bind data label number format to source cells Aspose.Cells | chart series data label formatting Aspose.Cells .NET | how to use DataLabels.NumberFormatLinked in Aspose.Cells
// Developer Intent: Enable DataLabels.NumberFormatLinked for the first chart series so that data label formatting is automatically taken from the linked source cells.
// Use Cases: Create charts where data labels display values with custom units (e.g., "100 units") stored in a separate column. | Maintain consistent label appearance when cell formatting changes, without manually updating each label. | Generate automated Excel reports with dynamic label formatting linked to source data.
// AI Prompts: Generate C# code using Aspose.Cells to link chart data label formatting to source cells with DataLabels.NumberFormatLinked. | Explain the impact of setting DataLabels.NumberFormatLinked to true on chart data labels in Aspose.Cells. | Show how updating the number format of linked cells automatically updates chart data labels when NumberFormatLinked is enabled.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsDemo
{
    // Demonstrates how to create a workbook, add a column chart, link data labels to formatted cells, and set DataLabels.NumberFormatLinked = true so the label appearance follows the source cell formatting. The example saves the result as an XLSX file.
    class Program
    {
        static void Main(string[] args)
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
            sheet.Cells["C1"].PutValue("Formatted Value");
            sheet.Cells["C2"].PutValue("100 units");
            sheet.Cells["C3"].PutValue("200 units");

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Define the series data range and category labels
            chart.NSeries.Add("B2:B3", true);
            chart.NSeries.CategoryData = "A2:A3";

            // Access the first series
            Series firstSeries = chart.NSeries[0];

            // Enable data labels and bind number format to the source cells
            firstSeries.DataLabels.ShowValue = true;
            firstSeries.DataLabels.LinkedSource = "C2:C3"; // link to formatted cells
            firstSeries.DataLabels.NumberFormatLinked = true; // bind formatting

            // Save the workbook to an XLSX file
            workbook.Save("DataLabelsNumberFormatLinkedDemo.xlsx");
        }
    }
}
