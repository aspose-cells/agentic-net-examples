// Title: Create a column chart with data labels sourced from a separate worksheet range using Aspose.Cells for .NET
// AI Prompts: Write C# code that adds a column chart to a worksheet and links its data labels to a cell range on another sheet with Aspose.Cells. | Show how to hide the default numeric values of chart data labels and display custom text strings from a different worksheet using Aspose.Cells. | Demonstrate configuring Series.DataLabels.LinkedSource to show custom label strings for a column chart in a .NET Excel workbook.
// Common Searches: Aspose.Cells custom data label text from another worksheet | C# link chart data labels to a range on a different sheet Aspose.Cells | display custom strings instead of values in column chart labels using Aspose.Cells for .NET | Series.DataLabels.ShowCellRange true example Aspose.Cells
// Tags: Aspose.Cells column chart custom data labels | Series.DataLabels.LinkedSource usage | chart data labels from external worksheet range | hide default label values Aspose.Cells | C# Excel chart custom label text

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a workbook with two worksheets: one holds chart data and the other stores custom label strings. It adds a column chart to the data sheet, configures the series to hide default values, enables cell‑range labels, and links the labels to the range Labels!A2:A4. The chart is saved as ChartWithCustomDataLabels.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the default worksheet (Sheet1)
        Workbook workbook = new Workbook();
        Worksheet sheet1 = workbook.Worksheets[0];
        sheet1.Name = "Data";

        // Populate chart data in Sheet1
        sheet1.Cells["A1"].PutValue("Category");
        sheet1.Cells["A2"].PutValue("A");
        sheet1.Cells["A3"].PutValue("B");
        sheet1.Cells["A4"].PutValue("C");

        sheet1.Cells["B1"].PutValue("Value");
        sheet1.Cells["B2"].PutValue(120);
        sheet1.Cells["B3"].PutValue(80);
        sheet1.Cells["B4"].PutValue(150);

        // Add a second worksheet to hold custom label texts
        Worksheet sheet2 = workbook.Worksheets.Add("Labels");
        sheet2.Cells["A1"].PutValue("CustomLabel");
        sheet2.Cells["A2"].PutValue("High");
        sheet2.Cells["A3"].PutValue("Medium");
        sheet2.Cells["A4"].PutValue("Low");

        // Add a column chart to Sheet1
        int chartIndex = sheet1.Charts.Add(ChartType.Column, 5, 0, 20, 12);
        Chart chart = sheet1.Charts[chartIndex];

        // Set the data range for the series and categories
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Enable data labels for the first series
        Series series = chart.NSeries[0];
        series.DataLabels.ShowValue = false;          // hide default value
        series.DataLabels.ShowCellRange = true;       // enable custom text from cells
        // Link the custom text range (Sheet2!A2:A4) to the data labels
        series.DataLabels.LinkedSource = "Labels!A2:A4";

        // Optional: adjust label appearance
        series.DataLabels.Position = LabelPositionType.InsideEnd;
        series.DataLabels.Font.Color = System.Drawing.Color.Blue;

        // Save the workbook
        workbook.Save("ChartWithCustomDataLabels.xlsx");
    }
}
