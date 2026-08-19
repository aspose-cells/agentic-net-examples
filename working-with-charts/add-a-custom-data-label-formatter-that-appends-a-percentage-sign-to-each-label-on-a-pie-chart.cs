// Title: C# – Append a Percent Sign to Pie Chart Data Labels Using Aspose.Cells
// Description: This example creates a workbook, adds sample categories and values, inserts a pie chart, binds the series to the data range, and configures the series’ data labels to display only percentages with a trailing “%” by setting the NumberFormat to "0%" and disabling raw values. The workbook is then saved as an Excel file.
// Keywords: Aspose.Cells | C# pie chart | data label formatting | percentage label | custom number format | Excel chart labels | Aspose.Cells .NET | chart data labels | show percentage | hide values
// Common Searches: Aspose.Cells show percentage on pie chart | C# set chart data label format Aspose.Cells | add % sign to pie chart labels .NET | custom data label formatter Aspose.Cells | how to hide values in chart labels Aspose.Cells
// Developer Intent: Configure pie chart data labels to show percentages with a % sign and omit raw values.
// Use Cases: Generate a pie chart where each slice label shows a formatted percentage | Create Excel reports with clean percentage-only labels on charts | Standardize chart label appearance across multiple workbooks using Aspose.Cells
// AI Prompts: Write C# code with Aspose.Cells that creates a pie chart and formats data labels to display only percentages followed by a percent sign. | Show how to set NumberFormat = "0%" for chart data labels and hide the underlying values in Aspose.Cells. | Explain the steps to enable percentage labels and apply a custom formatter to a pie chart using Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsCustomDataLabel
{
    // This example creates a workbook, adds sample categories and values, inserts a pie chart, binds the series to the data range, and configures the series’ data labels to display only percentages with a trailing “%” by setting the NumberFormat to "0%" and disabling raw values. The workbook is then saved as an Excel file.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the pie chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["A4"].PutValue("C");

            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);

            // Add a pie chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Pie, 5, 0, 20, 8);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the series
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Enable data labels for the first series
            DataLabels dataLabels = chart.NSeries[0].DataLabels;
            dataLabels.ShowPercentage = true;          // Show the percentage value
            dataLabels.ShowValue = false;              // Hide the raw value (optional)
            dataLabels.NumberFormat = "0%";            // Custom format that appends a percent sign

            // Save the workbook
            workbook.Save("PieChartCustomDataLabels.xlsx");
        }
    }
}
