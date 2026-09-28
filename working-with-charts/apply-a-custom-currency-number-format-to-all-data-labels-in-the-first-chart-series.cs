// Title: How to apply a custom currency number format to data labels of the first series in an Aspose.Cells column chart (C#)
// AI Prompts: Write C# code using Aspose.Cells to enable data labels on the first series of a column chart and set their number format to "$#,##0.00". | Show the steps for formatting chart data labels as currency in an Excel workbook created with Aspose.Cells for .NET.
// Common Searches: Aspose.Cells C# set currency format for chart series data labels | How to display data labels as $ currency in a column chart using Aspose.Cells | Apply custom number format to Excel chart data labels with Aspose.Cells .NET | Enable and format data labels on the first series of an Aspose.Cells chart | C# Aspose.Cells chart data label number format example
// Tags: Aspose.Cells chart series data label currency format | C# set data label number format Aspose.Cells | column chart data label customization Aspose.Cells | Excel chart currency formatting with Aspose.Cells | apply number format to chart labels .NET

using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a workbook, adds sample data, inserts a column chart, enables data labels for the first series, applies the currency format "$#,##0.00" to those labels, and saves the file as ChartDataLabelsCurrency.xlsx.
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

        // Set the data range for the chart
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Access the first series of the chart
        Series series = chart.NSeries[0];

        // Enable data labels for the series
        series.DataLabels.ShowValue = true;

        // Apply a custom currency number format to all data labels in this series
        series.DataLabels.NumberFormat = "$#,##0.00";

        // Save the workbook to a file
        workbook.Save("ChartDataLabelsCurrency.xlsx");
    }
}
