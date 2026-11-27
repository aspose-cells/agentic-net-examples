// Title: How to display custom text data labels from a worksheet range in an Aspose.Cells column chart (C#)
// Common Searches: Aspose.Cells C# set chart data labels to values from a cell range | link data labels to custom strings in another column using Aspose.Cells | hide numeric values and show custom labels in an Aspose.Cells column chart | use DataLabels.LinkedSource property in Aspose.Cells C# example | display custom data label text from worksheet cells in an Aspose.Cells chart
// Tags: Aspose.Cells chart custom data labels | DataLabels.LinkedSource C# | column chart data label customization Aspose.Cells | hide numeric values chart data labels Aspose.Cells | link chart labels to worksheet range Aspose.Cells

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a workbook, adds category, value, and custom label columns, inserts a column chart, assigns the value series and categories, then configures the series data labels to hide the default numeric values and display custom strings from the range C2:C4. Font color, size, and label position are customized before saving the workbook as an XLSX file.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet dataSheet = workbook.Worksheets[0];
        dataSheet.Name = "Data";

        // Populate sample data for the chart
        dataSheet.Cells["A1"].PutValue("Category");
        dataSheet.Cells["A2"].PutValue("A");
        dataSheet.Cells["A3"].PutValue("B");
        dataSheet.Cells["A4"].PutValue("C");

        dataSheet.Cells["B1"].PutValue("Value");
        dataSheet.Cells["B2"].PutValue(120);
        dataSheet.Cells["B3"].PutValue(80);
        dataSheet.Cells["B4"].PutValue(150);

        // Store custom label strings in a separate column (C)
        dataSheet.Cells["C1"].PutValue("CustomLabel");
        dataSheet.Cells["C2"].PutValue("Low");
        dataSheet.Cells["C3"].PutValue("Medium");
        dataSheet.Cells["C4"].PutValue("High");

        // Add a column chart to the worksheet
        int chartIndex = dataSheet.Charts.Add(ChartType.Column, 6, 0, 20, 10);
        Chart chart = dataSheet.Charts[chartIndex];
        chart.Title.Text = "Custom Data Labels Demo";

        // Set the series data range and category (X‑axis) data
        chart.NSeries.Add("B2:B4", true);          // Values
        chart.NSeries.CategoryData = "A2:A4";      // Categories

        // Configure data labels to display the custom text from C2:C4
        Series series = chart.NSeries[0];
        series.DataLabels.ShowValue = false;       // Hide the default numeric value
        series.DataLabels.ShowCellRange = true;    // Enable linking to a cell range
        series.DataLabels.LinkedSource = "C2:C4";  // Range containing custom strings
        series.DataLabels.Font.Color = Color.DarkBlue;
        series.DataLabels.Font.Size = 10;
        series.DataLabels.Position = LabelPositionType.OutsideEnd;

        // Save the workbook
        string outputPath = "CustomDataLabels.xlsx";
        workbook.Save(outputPath, SaveFormat.Xlsx);
        Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
    }
}
