// Title: C# Aspose.Cells – Disable Text Wrap for All Chart Data Labels Across Worksheets
// Description: Load a workbook, loop through every worksheet, each chart, and every series, then set DataLabels.IsTextWrapped = false to turn off text wrapping for all chart data labels before saving the file. This example shows a batch‑processing approach using Aspose.Cells for .NET.
// Keywords: Aspose.Cells chart data label wrap | disable text wrap chart labels .NET | batch update chart labels Aspose | iterate worksheets charts series C# | DataLabels.IsTextWrapped false
// Common Searches: how to turn off text wrap for chart data labels in Aspose.Cells | C# batch disable data label wrapping for all charts | Aspose.Cells iterate all worksheets and charts | set chart data label wrap property programmatically | disable chart label wrap across multiple sheets
// Developer Intent: Turn off text wrapping for every chart series data label in all worksheets of an Excel file.
// Use Cases: Standardize report visuals so labels never wrap, improving on‑screen clarity. | Prepare workbooks for printing where wrapped labels cause layout issues. | Automate bulk cleanup of chart formatting before sharing Excel files with stakeholders.
// AI Prompts: Generate C# code using Aspose.Cells that disables text wrapping for data labels in every chart across all worksheets and saves the workbook. | Explain how to traverse worksheets, charts, and series to set DataLabels.IsTextWrapped = false, then export the modified file. | Show how to modify the loop to apply the wrap‑disable setting only to column and bar charts while leaving other chart types unchanged.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Load a workbook, loop through every worksheet, each chart, and every series, then set DataLabels.IsTextWrapped = false to turn off text wrapping for all chart data labels before saving the file. This example shows a batch‑processing approach using Aspose.Cells for .NET.
class BatchDisableDataLabelWrap
{
    static void Main()
    {
        // Load an existing workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Iterate through all worksheets in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Iterate through all charts in the current worksheet
            foreach (Chart chart in sheet.Charts)
            {
                // Iterate through all series of the chart
                foreach (Series series in chart.NSeries)
                {
                    // Access the data labels of the series
                    DataLabels labels = series.DataLabels;

                    // Disable text wrapping for the data labels
                    labels.IsTextWrapped = false;
                }
            }
        }

        // Save the modified workbook (replace with your desired output path)
        workbook.Save("output.xlsx");
    }
}
