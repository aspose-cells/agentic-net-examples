// Title: Set Currency Number Format for Chart Series Data Labels with Aspose.Cells for .NET
// Description: Shows how to create a workbook, add a column chart, enable data labels on the first series, and apply the "$#,##0.00" currency format to every label using Aspose.Cells for .NET before saving the file.
// Keywords: Aspose.Cells C# chart data label format | currency number format Aspose.Cells | format chart series labels .NET | Excel chart data label styling | apply number format to Aspose chart | chart data labels custom format | Aspose.Cells chart formatting example
// Common Searches: Aspose.Cells how to format chart data labels as currency | C# set number format for chart series labels | apply custom number format to Excel chart using Aspose | chart data label currency format Aspose.Cells | format first series data labels in column chart .NET
// Developer Intent: Display the first series' data labels in a chart using a specific currency format.
// Use Cases: Financial dashboards where column values need to appear as dollars with two decimal places. | Automated sales reports that export Excel charts with currency‑formatted data labels for quick review. | Generating investor presentations where chart labels must follow a consistent monetary style.
// AI Prompts: Provide C# code that sets the "$#,##0.00" currency format for all data labels of the first series in an Aspose.Cells chart. | Show an example of enabling data labels on a column chart and applying a custom number format using Aspose.Cells for .NET.

using Aspose.Cells;
using Aspose.Cells.Charts;

// Shows how to create a workbook, add a column chart, enable data labels on the first series, and apply the "$#,##0.00" currency format to every label using Aspose.Cells for .NET before saving the file.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate sample data for the chart
        worksheet.Cells["A1"].PutValue("Category");
        worksheet.Cells["A2"].PutValue("A");
        worksheet.Cells["A3"].PutValue("B");
        worksheet.Cells["A4"].PutValue("C");
        worksheet.Cells["B1"].PutValue("Value");
        worksheet.Cells["B2"].PutValue(1000);
        worksheet.Cells["B3"].PutValue(2000);
        worksheet.Cells["B4"].PutValue(3000);

        // Add a column chart to the worksheet
        int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
        Chart chart = worksheet.Charts[chartIndex];

        // Set the data range for the series and categories
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Access the first series
        Series series = chart.NSeries[0];

        // Enable data labels for the series
        series.DataLabels.ShowValue = true;

        // Apply a custom currency number format to all data labels in this series
        series.DataLabels.NumberFormat = "$#,##0.00";

        // Save the workbook to a file
        workbook.Save("ChartDataLabelsCurrency.xlsx");
    }
}
