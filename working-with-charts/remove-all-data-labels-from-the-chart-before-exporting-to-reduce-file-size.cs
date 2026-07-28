// Title: Remove Data Labels from an Aspose.Cells Chart (.NET) to Reduce Workbook Size
// Description: Creates a workbook, adds sample data, inserts a column chart, enables data labels for demonstration, then deletes all series‑ and point‑level data labels using the IsDeleted flag before saving the file as XLSX.
// Keywords: Aspose.Cells remove chart data labels | delete chart labels .NET | Aspose.Cells IsDeleted property | reduce Excel file size chart labels | clear data labels Aspose.Cells
// Common Searches: how to delete data labels from a chart using Aspose.Cells for .NET | Aspose.Cells remove point data labels to shrink workbook | chart IsDeleted property Aspose.Cells example | disable chart data labels before saving Excel with Aspose.Cells
// Developer Intent: Eliminate all data labels from a chart so the generated Excel file is smaller and visually cleaner.
// Use Cases: Export a sales column chart without labels to keep the XLSX lightweight. | Prepare a dashboard workbook where charts are label‑free for a cleaner look. | Programmatically clean existing template charts by removing series and point labels before distribution.
// AI Prompts: Generate C# code with Aspose.Cells that removes all data labels from every series in a chart and saves the workbook. | Explain the effect of the IsDeleted flag on series and point data labels in Aspose.Cells charts. | Suggest additional techniques to minimize Excel file size when charts contain data labels using Aspose.Cells.

using Aspose.Cells;
using Aspose.Cells.Charts;

// Creates a workbook, adds sample data, inserts a column chart, enables data labels for demonstration, then deletes all series‑ and point‑level data labels using the IsDeleted flag before saving the file as XLSX.
class RemoveDataLabelsDemo
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
        worksheet.Cells["B2"].PutValue(10);
        worksheet.Cells["B3"].PutValue(20);
        worksheet.Cells["B4"].PutValue(30);

        // Add a column chart to the worksheet
        int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
        Chart chart = worksheet.Charts[chartIndex];
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // (Optional) Enable data labels to demonstrate removal
        chart.NSeries[0].DataLabels.ShowValue = true;

        // Remove all data labels from the chart
        foreach (Series series in chart.NSeries)
        {
            // Mark the series-level data labels as deleted
            series.DataLabels.IsDeleted = true;

            // Also ensure any point-level data labels are deleted
            foreach (ChartPoint point in series.Points)
            {
                point.DataLabels.IsDeleted = true;
            }
        }

        // Save the workbook with the chart that no longer contains data labels
        workbook.Save("ChartWithoutDataLabels.xlsx", SaveFormat.Xlsx);
    }
}
