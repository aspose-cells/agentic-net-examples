// Title: Aspose.Cells for .NET – Disable Text Wrap on All Chart Data Labels Across Worksheets
// Description: Loads a workbook, iterates through every worksheet and chart, and sets both series‑level and point‑level DataLabels.IsTextWrapped to false, then saves the file. Ideal for batch‑updating chart label formatting in Excel with Aspose.Cells.
// Keywords: Aspose.Cells disable chart label wrap | C# chart data label no wrap | batch update chart labels Aspose | iterate worksheets charts Aspose.Cells | DataLabels.IsTextWrapped false
// Common Searches: how to turn off text wrap for chart data labels in Aspose.Cells | disable data label wrapping for all charts in a workbook .NET | loop through worksheets and charts to change label properties Aspose | remove text wrapping from Excel chart labels programmatically
// Developer Intent: Turn off text wrapping for every chart data label in all worksheets of an Excel workbook.
// Use Cases: Prepare printable reports where chart labels must stay on a single line. | Standardize chart appearance across multiple sheets before PDF conversion. | Clean up legacy workbooks that have wrapped labels causing layout problems.
// AI Prompts: Write C# code using Aspose.Cells to disable text wrap on chart data labels in every worksheet and save the workbook. | Show how to also change the font size and color of data labels after disabling wrap for each series and point. | Explain how to limit the operation to only ColumnClustered charts while leaving other chart types unchanged.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Loads a workbook, iterates through every worksheet and chart, and sets both series‑level and point‑level DataLabels.IsTextWrapped to false, then saves the file. Ideal for batch‑updating chart label formatting in Excel with Aspose.Cells.
class Program
{
    static void Main()
    {
        // Load an existing workbook (replace with your file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Loop through all worksheets in the workbook
        foreach (Worksheet worksheet in workbook.Worksheets)
        {
            // Loop through all charts on the current worksheet
            foreach (Chart chart in worksheet.Charts)
            {
                // Loop through each series in the chart
                foreach (Series series in chart.NSeries)
                {
                    // Disable text wrapping for the series data labels
                    series.DataLabels.IsTextWrapped = false;

                    // Also disable text wrapping for each individual point's data label
                    foreach (ChartPoint point in series.Points)
                    {
                        point.DataLabels.IsTextWrapped = false;
                    }
                }
            }
        }

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }
}
