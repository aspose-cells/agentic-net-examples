// Title: C# – Remove All Chart Data Labels with AspNet.Cells to Reduce Workbook Size
// Description: Demonstrates how to create a workbook, add a column chart, enable data labels for illustration, then delete every label by setting Series.DataLabels.IsDeleted = true, and finally save the file as a compact XLSX.
// Keywords: Aspose.Cells remove chart data labels | C# delete chart data labels | Aspose.Cells reduce Excel file size | Series.DataLabels.IsDeleted | chart export optimization .NET
// Common Searches: how to delete data labels from a chart using Aspose.Cells C# | remove chart labels to shrink workbook size | Aspose.Cells Series.DataLabels.IsDeleted example | C# chart data labels removal Aspose | optimize Excel file size by removing chart labels
// Developer Intent: Eliminate every data label from a chart before saving to minimize the workbook’s size.
// Use Cases: Generate a clean chart for client distribution without visible values. | Iterate through all series in a multi‑series chart and turn off labels to meet reporting standards. | Compress an Excel file containing many charts by removing unnecessary label data.
// AI Prompts: Write C# code that removes data labels from all series in an existing Aspose.Cells chart while preserving other formatting. | Explain the effect of setting Series.DataLabels.IsDeleted = true on the saved XLSX file size and when this technique is appropriate. | Show how to selectively delete data labels from specific series in a chart using Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // Demonstrates how to create a workbook, add a column chart, enable data labels for illustration, then delete every label by setting Series.DataLabels.IsDeleted = true, and finally save the file as a compact XLSX.
    class RemoveDataLabelsDemo
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["A4"].PutValue("C");

            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);

            // Add a column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the chart
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Enable data labels initially (optional, just to demonstrate removal)
            chart.NSeries[0].DataLabels.ShowValue = true;

            // Remove all data labels from every series in the chart
            foreach (Series series in chart.NSeries)
            {
                // Mark the DataLabels object as deleted – this removes the labels completely
                series.DataLabels.IsDeleted = true;
            }

            // Save the workbook with the chart (data labels removed)
            workbook.Save("ChartWithoutDataLabels.xlsx", SaveFormat.Xlsx);
        }
    }
}
