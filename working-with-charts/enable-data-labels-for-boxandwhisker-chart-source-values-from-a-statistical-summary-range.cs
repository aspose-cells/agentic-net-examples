// Title: Aspose.Cells for .NET – Add Data Labels to a Box‑and‑Whisker Chart from a Median Range
// Description: Demonstrates how to create a workbook, populate category, raw values and median cells, insert a Box‑and‑Whisker chart, enable data labels, link them to a summary range (C2:C5), set label position to center, and apply a two‑decimal number format before saving the file.
// Keywords: Aspose.Cells | C# | .NET | Box‑and‑Whisker chart | data labels | linked cell range | median values | chart customization | LabelPositionType.Center | NumberFormat 0.00
// Common Searches: Aspose.Cells box plot data labels linked to cells | How to show median values on a Box‑and‑Whisker chart in .NET | Enable and format data labels for Aspose.Cells chart | Link chart data labels to a summary range using Aspose.Cells
// Developer Intent: Enable data labels on a Box‑and‑Whisker chart and bind them to a separate median range for custom display.
// Use Cases: Financial reports that need quarterly medians displayed on each box. | Quality‑control dashboards where statistical summaries are shown as chart labels. | Automating existing workbooks to add centered, formatted median labels to box‑plot series.
// AI Prompts: Write C# code with Aspose.Cells to create a Box‑and‑Whisker chart, show data labels, and link them to a median column formatted to two decimals. | Explain step‑by‑step how to bind data labels to a cell range for a Box‑and‑Whisker series and adjust label position and number format in Aspose.Cells. | Provide instructions to modify a workbook so that each box in a Box‑and‑Whisker chart displays its median from cells C2:C5 as a centered label.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Demonstrates how to create a workbook, populate category, raw values and median cells, insert a Box‑and‑Whisker chart, enable data labels, link them to a summary range (C2:C5), set label position to center, and apply a two‑decimal number format before saving the file.
class BoxWhiskerDataLabelsDemo
{
    static void Main()
    {
        // Create a new workbook (create rule)
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // -------------------------------------------------
        // Prepare source data for the box‑and‑whisker chart
        // -------------------------------------------------
        // Category column
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["A2"].PutValue("Q1");
        sheet.Cells["A3"].PutValue("Q1");
        sheet.Cells["A4"].PutValue("Q1");
        sheet.Cells["A5"].PutValue("Q2");
        sheet.Cells["A6"].PutValue("Q2");
        sheet.Cells["A7"].PutValue("Q2");

        // Raw values (will be used by the chart)
        sheet.Cells["B1"].PutValue("Values");
        sheet.Cells["B2"].PutValue(15);
        sheet.Cells["B3"].PutValue(25);
        sheet.Cells["B4"].PutValue(30);
        sheet.Cells["B5"].PutValue(20);
        sheet.Cells["B6"].PutValue(35);
        sheet.Cells["B7"].PutValue(40);

        // Statistical summary that will be shown in data labels
        // (e.g., median values for each category)
        sheet.Cells["C1"].PutValue("Median");
        sheet.Cells["C2"].PutValue(25); // median for Q1
        sheet.Cells["C5"].PutValue(30); // median for Q2

        // -------------------------------------------------
        // Add a Box‑and‑Whisker chart
        // -------------------------------------------------
        int chartIndex = sheet.Charts.Add(ChartType.BoxWhisker, 9, 0, 25, 15);
        Chart chart = sheet.Charts[chartIndex];

        // Set the data range for the series (values only)
        chart.NSeries.Add("B2:B7", true);
        // Set the category (X‑axis) data
        chart.NSeries.CategoryData = "A2:A7";

        // -------------------------------------------------
        // Enable data labels and bind them to the summary range
        // -------------------------------------------------
        Series series = chart.NSeries[0];

        // Show the value (the box‑plot statistic) on the label
        series.DataLabels.ShowValue = true;

        // Show the cell range (the median values) on the label
        series.DataLabels.ShowCellRange = true;
        // Link the label to the summary cells (C2:C5 contains the medians)
        series.DataLabels.LinkedSource = "C2:C5";

        // Optional: customize label appearance
        series.DataLabels.Position = LabelPositionType.Center;
        series.DataLabels.NumberFormat = "0.00";

        // -------------------------------------------------
        // Save the workbook (save rule)
        // -------------------------------------------------
        workbook.Save("BoxWhisker_WithDataLabels.xlsx");
    }
}
