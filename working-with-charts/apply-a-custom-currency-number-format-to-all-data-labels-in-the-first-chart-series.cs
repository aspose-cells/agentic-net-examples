// Title: Apply a Custom Currency Format to the First Chart Series Data Labels with Aspose.Cells for .NET
// Description: This C# example creates a workbook, inserts sample categories and values, adds a column chart, enables data labels for the first series, and sets the series' DataLabels.NumberFormat to the currency pattern "$#,##0.00" before saving the file as an Excel workbook.
// Keywords: Aspose.Cells | .NET chart formatting | C# chart data labels | currency number format | NumberFormat property | column chart Aspose | Excel financial chart | custom number format Excel | chart series labeling | Aspose.Cells tutorial
// Common Searches: Aspose.Cells set chart data label currency format | C# apply $#,##0.00 to chart series labels | How to format Excel chart data labels as dollars using Aspose | NumberFormat for chart data labels in Aspose.Cells .NET | Add currency formatting to first series of a column chart
// Developer Intent: Show the first series of a chart with data labels formatted as a custom currency string.
// Use Cases: Generate a sales dashboard where each column displays its value in dollars. | Create a financial report that requires accounting‑style currency labels on chart bars. | Export an Excel workbook for presentations with chart labels automatically formatted as $1,234.00.
// AI Prompts: Provide C# code using Aspose.Cells to add a column chart, enable data labels, and apply the "$#,##0.00" currency format to the first series. | Explain how to use the DataLabels.NumberFormat property to format chart series labels as currency in Aspose.Cells for .NET. | Show step‑by‑step instructions for creating a workbook, populating data, inserting a chart, and setting a custom number format on its data labels.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsCustomCurrencyDataLabels
{
    // This C# example creates a workbook, inserts sample categories and values, adds a column chart, enables data labels for the first series, and sets the series' DataLabels.NumberFormat to the currency pattern "$#,##0.00" before saving the file as an Excel workbook.
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
            sheet.Cells["A4"].PutValue("C");

            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["B2"].PutValue(1000);
            sheet.Cells["B3"].PutValue(2000);
            sheet.Cells["B4"].PutValue(3000);

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the series and categories
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Access the first series
            Series firstSeries = chart.NSeries[0];

            // Enable data labels for the series
            firstSeries.DataLabels.ShowValue = true;

            // Apply a custom currency number format to all data labels in this series
            // Example format: $1,234.00
            firstSeries.DataLabels.NumberFormat = "$#,##0.00";

            // Save the workbook to a file
            workbook.Save("ChartWithCurrencyDataLabels.xlsx");
        }
    }
}
