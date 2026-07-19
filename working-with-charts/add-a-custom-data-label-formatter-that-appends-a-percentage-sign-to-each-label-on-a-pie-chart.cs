// Title: Format Pie Chart Data Labels as Percentages with Aspose.Cells for .NET (C#)
// Description: Creates a workbook, adds sample categories and values, inserts a pie chart, and configures the chart's data labels to show only calculated percentages using the custom number format "0%" while hiding raw values. The workbook is saved as PieChartCustomDataLabel.xlsx.
// Keywords: Aspose.Cells pie chart percentage labels | C# chart data label formatter | Aspose.Cells custom number format | show percentage data labels Aspose.Cells | .NET Excel chart label styling
// Common Searches: Aspose.Cells display % on pie chart labels | C# set custom number format for chart data labels | How to hide values and show percentages in Aspose.Cells chart | Add percent sign to Excel pie chart labels using Aspose
// Developer Intent: Apply a custom formatter so each pie‑chart slice label displays a calculated percentage followed by a % sign.
// Use Cases: Generate Excel reports where pie‑chart labels show only percentages for clearer visual analysis. | Standardize label appearance across multiple charts by applying a "0%" number format. | Create dashboards that hide raw values and present concise percentage information on pie slices.
// AI Prompts: Write C# code with Aspose.Cells that builds a pie chart and formats its data labels to display percentages with a trailing % sign. | Show how to enable DataLabels.ShowPercentage and set NumberFormat to "0%" while disabling raw values in an Aspose.Cells chart. | Explain the steps to hide numeric values and show only formatted percentages on a pie chart using Aspose.Cells for .NET.

using Aspose.Cells;
using Aspose.Cells.Charts;

// Creates a workbook, adds sample categories and values, inserts a pie chart, and configures the chart's data labels to show only calculated percentages using the custom number format "0%" while hiding raw values. The workbook is saved as PieChartCustomDataLabel.xlsx.
class Program
{
    static void Main()
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

        // Set the data range for the chart
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Enable data labels and apply a custom number format that appends a percentage sign
        DataLabels dataLabels = chart.NSeries[0].DataLabels;
        dataLabels.ShowPercentage = true;   // calculate percentage values
        dataLabels.ShowValue = false;       // hide raw values
        dataLabels.NumberFormat = "0%";     // custom formatter to display as percentage with % sign

        // Save the workbook
        workbook.Save("PieChartCustomDataLabel.xlsx");
    }
}
