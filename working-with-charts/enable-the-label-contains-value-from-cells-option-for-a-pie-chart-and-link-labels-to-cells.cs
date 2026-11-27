// Title: How to enable "Label Contains – Value From Cells" for a pie chart and link custom labels to cells using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# set pie chart data labels to use values from a cell range | link custom label cells to a pie chart series with Aspose.Cells .NET | disable numeric values on pie chart data labels in Aspose.Cells | show cell range as data label for a pie chart using Aspose.Cells API | set label position to OutsideEnd for pie chart labels in Aspose.Cells C#
// Tags: cell range data labels Aspose.Cells | ShowCellRange series property Aspose.Cells | custom label cells linking Aspose.Cells chart | disable numeric data label Aspose.Cells | outside end label position Aspose.Cells | pie chart NSeries values range Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using System.Drawing;

// The example creates a workbook, fills columns with categories, numeric values, and custom label strings, adds a pie chart, configures the series to use the value range and category labels, disables the default numeric label, enables ShowCellRange, links the data labels to cells C2:C4, customizes font color, size, and positions the labels outside the slices, then saves the file as an XLSX workbook.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // -------------------------------------------------
        // 1. Populate source data for the pie chart
        // -------------------------------------------------
        // Category names
        worksheet.Cells["A1"].PutValue("Category");
        worksheet.Cells["A2"].PutValue("Apple");
        worksheet.Cells["A3"].PutValue("Banana");
        worksheet.Cells["A4"].PutValue("Cherry");

        // Numeric values for the pie slices
        worksheet.Cells["B1"].PutValue("Value");
        worksheet.Cells["B2"].PutValue(30);
        worksheet.Cells["B3"].PutValue(45);
        worksheet.Cells["B4"].PutValue(25);

        // Custom labels that will be shown on the chart (value from cells)
        worksheet.Cells["C1"].PutValue("Label");
        worksheet.Cells["C2"].PutValue("30 pcs");
        worksheet.Cells["C3"].PutValue("45 pcs");
        worksheet.Cells["C4"].PutValue("25 pcs");

        // -------------------------------------------------
        // 2. Add a Pie chart to the worksheet
        // -------------------------------------------------
        // Parameters: chart type, upper‑left row, upper‑left column, lower‑right row, lower‑right column
        int chartIndex = worksheet.Charts.Add(ChartType.Pie, 6, 0, 20, 10);
        Chart pieChart = worksheet.Charts[chartIndex];

        // Set the data range for the series (values) and the category axis (labels)
        pieChart.NSeries.Add("B2:B4", true);          // Values
        pieChart.NSeries.CategoryData = "A2:A4";     // Categories

        // -------------------------------------------------
        // 3. Configure Data Labels to use cell values
        // -------------------------------------------------
        Series series = pieChart.NSeries[0];

        // Disable the default numeric value that Excel would otherwise display
        series.DataLabels.ShowValue = false;

        // Enable the "Label contains – Value from cells" option
        series.DataLabels.ShowCellRange = true;

        // Link the data labels to the custom label cells
        series.DataLabels.LinkedSource = "C2:C4";

        // Optional: style the labels
        series.DataLabels.Font.Color = Color.DarkBlue;
        series.DataLabels.Font.Size = 10;
        series.DataLabels.Position = LabelPositionType.OutsideEnd;

        // -------------------------------------------------
        // 4. Save the workbook
        // -------------------------------------------------
        string outputPath = "PieChartLabelFromCells.xlsx";
        workbook.Save(outputPath, SaveFormat.Xlsx);
        Console.WriteLine($"Workbook saved successfully to \"{outputPath}\"");
    }
}
