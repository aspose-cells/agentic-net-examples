// Title: Aspose.Cells C# – Remove Chart Data Labels to Reduce File Size
// Description: Demonstrates how to delete all data labels from every series in an Aspose.Cells chart using the Series.DataLabels.IsDeleted property, then save the workbook as an XLSX file. Removing labels shrinks the generated Excel file and cleans up the visual output.
// Keywords: Aspose.Cells chart data labels | C# delete chart labels | reduce Excel file size | Series.DataLabels.IsDeleted | .NET chart cleanup | Aspose.Cells export optimization
// Common Searches: how to hide data labels Aspose.Cells C# | remove chart labels before saving workbook | Aspose.Cells reduce XLSX size by deleting labels | C# chart data labels IsDeleted example | Aspose.Cells chart cleanup for distribution
// Developer Intent: Programmatically eliminate all data labels from an Aspose.Cells chart prior to saving the workbook.
// Use Cases: Prepare a report workbook for external distribution without unnecessary label clutter. | Batch‑process multiple worksheets to minimize Excel file size for web downloads. | Generate clean visual charts for dashboards where data values are shown elsewhere.
// AI Prompts: Write C# code that uses Aspose.Cells to remove data labels from every series in a chart and then saves the file. | Explain the effect of the Series.DataLabels.IsDeleted flag on Excel file size and rendering. | Show how to conditionally delete data labels only when a chart exceeds a certain number of data points.

using Aspose.Cells;
using Aspose.Cells.Charts;

// Demonstrates how to delete all data labels from every series in an Aspose.Cells chart using the Series.DataLabels.IsDeleted property, then save the workbook as an XLSX file. Removing labels shrinks the generated Excel file and cleans up the visual output.
class RemoveDataLabels
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
        chart.NSeries.Add("B2:B4", true);          // Set values
        chart.NSeries.CategoryData = "A2:A4";      // Set categories

        // (Optional) Enable data labels to demonstrate removal
        chart.NSeries[0].DataLabels.ShowValue = true;

        // Remove all data labels from every series in the chart
        foreach (Series series in chart.NSeries)
        {
            series.DataLabels.IsDeleted = true;   // Marks the data labels as deleted
        }

        // Save the workbook with the chart that no longer contains data labels
        workbook.Save("ChartWithoutDataLabels.xlsx", SaveFormat.Xlsx);
    }
}
