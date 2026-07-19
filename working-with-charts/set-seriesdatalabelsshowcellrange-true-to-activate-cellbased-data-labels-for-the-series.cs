// Title: C# – Enable cell‑based data labels for a chart series with Series.DataLabels.ShowCellRange in Aspose.Cells for .NET
// Description: This example creates a workbook, adds a column chart, populates category, value and label cells, then activates cell‑based data labels by setting Series.DataLabels.ShowCellRange to true and linking them to the range C2:C3 via Series.DataLabels.LinkedSource. The workbook is saved as SeriesDataLabelsShowCellRange.xlsx.
// Keywords: Aspose.Cells | Series.DataLabels.ShowCellRange | cell based data labels | chart series labels | C# Aspose.Cells chart | LinkedSource property | .NET Excel automation | column chart data labels | Excel cell linked labels | Aspose.Cells example
// Common Searches: Aspose.Cells showcellrange chart series C# | How to link chart data labels to cells in Aspose.Cells | Enable cell based data labels Aspose.Cells .NET | Series.DataLabels.ShowCellRange example | Aspose.Cells chart label from cell range
// Developer Intent: Activate cell‑based data labels for a chart series and bind them to a worksheet range.
// Use Cases: Show custom text from cells C2:C3 as data labels while also displaying the numeric values. | Create charts that automatically update labels when the linked worksheet cells change. | Combine numeric values and descriptive text on each data point by using ShowValue and ShowCellRange together.
// AI Prompts: Write C# code using Aspose.Cells to create a line chart where data labels are sourced from a specified cell range via ShowCellRange. | Explain how Series.DataLabels.ShowCellRange interacts with Series.DataLabels.LinkedSource when configuring chart labels in Aspose.Cells. | Provide step‑by‑step guidance to modify an existing workbook so multiple chart series each use different cell ranges for their data labels.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// This example creates a workbook, adds a column chart, populates category, value and label cells, then activates cell‑based data labels by setting Series.DataLabels.ShowCellRange to true and linking them to the range C2:C3 via Series.DataLabels.LinkedSource. The workbook is saved as SeriesDataLabelsShowCellRange.xlsx.
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
        sheet.Cells["C2"].PutValue("100 units");   // Custom label for first point
        sheet.Cells["C3"].PutValue("200 units");   // Custom label for second point

        // Add a column chart to the worksheet
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
        Chart chart = sheet.Charts[chartIndex];

        // Define the data range for the series and its categories
        chart.NSeries.Add("B2:B3", true);          // Values
        chart.NSeries.CategoryData = "A2:A3";      // Categories

        // Access the first series in the chart
        Series series = chart.NSeries[0];

        // Enable data labels and activate cell‑based data labels
        series.DataLabels.ShowValue = true;        // Show the numeric value
        series.DataLabels.ShowCellRange = true;    // Use cell range as data labels
        series.DataLabels.LinkedSource = "C2:C3";  // Link to cells containing custom text

        // Save the workbook to a file
        workbook.Save("SeriesDataLabelsShowCellRange.xlsx");
    }
}
