// Title: Aspose.Cells for .NET – Show Value and Series Name in Chart Data Labels with a Hyphen
// Description: Creates a workbook, adds category and numeric data, inserts a column chart, and configures the first series' DataLabels to display both the cell value and the series name using a custom " - " separator. The file is saved as an XLSX workbook.
// Keywords: Aspose.Cells chart data labels | custom separator DataLabels | show series name and value | C# column chart label formatting | Aspose.Cells DataLabelsSeparatorType | hyphen separator chart labels | .NET Excel chart customization
// Common Searches: Aspose.Cells show series name and value in chart label | custom data label separator Aspose.Cells .NET | column chart data labels hyphen separator C# | how to combine series name and value in Aspose.Cells chart | set DataLabelsSeparatorType to custom in Aspose.Cells
// Developer Intent: Add a column chart and configure its data labels to concatenate the series name and cell value with a hyphen.
// Use Cases: Sales dashboard where each column reads "Product – 1500" by merging product name and sales figure. | Financial report chart that displays "Revenue – $10M" in each bar for quick insight. | Project tracking chart showing "Task A – 75%" to combine task label and completion percentage.
// AI Prompts: Generate C# code with Aspose.Cells to create a line chart whose data labels combine series name and value using a slash separator. | Explain how to change the DataLabels separator from a hyphen to a pipe character for an existing Aspose.Cells chart. | Provide steps to enable ShowSeriesName and ShowValue for multiple series in a stacked column chart using Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsDataLabelsExample
{
    // Creates a workbook, adds category and numeric data, inserts a column chart, and configures the first series' DataLabels to display both the cell value and the series name using a custom " - " separator. The file is saved as an XLSX workbook.
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
            sheet.Cells["A4"].PutValue("C");

            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the series and categories
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Configure data labels to show both value and series name,
            // using a custom hyphen separator
            DataLabels dataLabels = chart.NSeries[0].DataLabels;
            dataLabels.ShowValue = true;                     // Show cell values
            dataLabels.ShowSeriesName = true;                // Show series name
            dataLabels.SeparatorType = DataLabelsSeparatorType.Custom; // Use custom separator
            dataLabels.SeparatorValue = " - ";               // Hyphen with spaces as separator

            // Save the workbook to an XLSX file
            workbook.Save("DataLabelsSeriesValueHyphen.xlsx");
        }
    }
}
