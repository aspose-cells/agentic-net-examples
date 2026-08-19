// Title: C# – Disable Text Wrapping for All Chart Data Labels in Every Worksheet with Aspose.Cells
// Description: Loads a workbook, iterates through each worksheet, chart, and series, sets DataLabels.IsTextWrapped to false, and saves the file, removing line‑breaks from every chart label in the workbook.
// Keywords: Aspose.Cells C# chart data label wrap | disable text wrap chart labels Aspose | batch update chart labels .NET | iterate worksheets charts Aspose.Cells | DataLabels.IsTextWrapped false
// Common Searches: how to turn off text wrap for chart labels in all worksheets asp.net | batch disable data label wrapping Aspose.Cells | C# loop through charts and series to change label properties | remove chart label wrapping from entire workbook
// Developer Intent: Turn off text wrapping for every chart data label across all worksheets in an Excel file using Aspose.Cells for .NET.
// Use Cases: Standardize label appearance in multi‑sheet financial reports. | Prevent layout shifts in automated dashboards where wrapped labels break design. | Prepare workbooks for distribution with consistent single‑line chart labels.
// AI Prompts: Generate C# code that disables text wrapping for chart data labels in all worksheets using Aspose.Cells. | Show how to change additional label properties (font size, color) while iterating through charts with Aspose.Cells. | Provide a snippet that applies IsTextWrapped = false only to column and bar charts across a workbook.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Loads a workbook, iterates through each worksheet, chart, and series, sets DataLabels.IsTextWrapped to false, and saves the file, removing line‑breaks from every chart label in the workbook.
class Program
{
    static void Main()
    {
        // Load an existing workbook (replace with your actual file path)
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
                    // Access the data labels of the series
                    DataLabels dataLabels = series.DataLabels;

                    // Disable text wrapping for the data labels
                    dataLabels.IsTextWrapped = false;
                }
            }
        }

        // Save the modified workbook (replace with your desired output path)
        workbook.Save("output.xlsx");
    }
}
