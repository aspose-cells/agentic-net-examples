// Title: Disable text wrapping for all chart data labels in every worksheet of an Excel workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code with Aspose.Cells that opens an Excel file, iterates through each worksheet and each chart, and sets DataLabels.IsTextWrapped = false for every series. | Show how to batch process an Excel workbook to turn off text wrap on chart data labels across all sheets and save the modified file. | Provide a step‑by‑step example of disabling data label text wrapping for multiple charts in a multi‑sheet workbook using Aspose.Cells.
// Common Searches: aspnet aspocells disable chart data label wrap for all worksheets | c# loop through worksheets and charts to turn off data label text wrapping | batch modify chart data labels in Excel using Aspose.Cells .NET | how to set IsTextWrapped false for chart series data labels in a workbook | remove text wrapping from chart labels in every sheet with Aspose.Cells
// Tags: chart data label wrap property Aspose.Cells | batch update chart settings across worksheets C# | chart data label wrap flag Aspose.Cells | traverse workbook charts Aspose.Cells | programmatic Excel chart label adjustment

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// // Loads an Excel workbook, iterates through each worksheet, each chart, and each series, disables text wrapping for the series' data labels, and saves the updated workbook.
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
                    // Access the data labels of the series
                    DataLabels dataLabels = series.DataLabels;

                    // Disable text wrapping for the data labels
                    dataLabels.IsTextWrapped = false;
                }
            }
        }

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }
}
