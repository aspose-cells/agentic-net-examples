// Title: C# – Apply Currency Number Format to First Chart Series Data Labels with Aspose.Cells
// Description: Creates a workbook, adds sample data, inserts a column chart, enables data labels for the first series, and sets the series' DataLabels.NumberFormat to "$#,##0.00" so every label displays values as US dollars with two decimal places. Saves the file as ChartWithCurrencyDataLabels.xlsx.
// Keywords: Aspose.Cells C# chart data labels | currency format Excel chart | NumberFormat property Aspose.Cells | .NET chart series formatting | Excel column chart label styling
// Common Searches: Aspose.Cells set data label format to currency | C# chart series number format Aspose.Cells | How to show dollar values on Excel chart labels using .NET | Apply custom number format to first series data labels
// Developer Intent: Apply a US‑dollar number format to all data labels of the first series in an Aspose.Cells chart.
// Use Cases: Financial reports that require chart labels to show amounts with a dollar sign and two decimals. | Automated sales dashboards where chart data labels must follow a consistent currency style. | Exporting Excel workbooks for accounting teams that need monetary values clearly formatted on charts.
// AI Prompts: Generate C# code that sets the NumberFormat of data labels for the first series of an Aspose.Cells chart to "$#,##0.00". | Show how to apply a custom currency format to multiple chart series data labels in an Aspose.Cells workbook. | Explain step‑by‑step how to enable and format data labels on an Excel chart using Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Creates a workbook, adds sample data, inserts a column chart, enables data labels for the first series, and sets the series' DataLabels.NumberFormat to "$#,##0.00" so every label displays values as US dollars with two decimal places. Saves the file as ChartWithCurrencyDataLabels.xlsx.
class ApplyCurrencyFormatToDataLabels
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
        Series series = chart.NSeries[0];

        // Enable data labels for the series
        series.DataLabels.ShowValue = true;

        // Apply a custom currency number format to all data labels in this series
        series.DataLabels.NumberFormat = "$#,##0.00";

        // Save the workbook
        workbook.Save("ChartWithCurrencyDataLabels.xlsx");
    }
}
