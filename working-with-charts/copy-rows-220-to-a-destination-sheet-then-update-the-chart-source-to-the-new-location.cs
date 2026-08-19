// Title: C# – Copy Rows 2‑20 to a New Worksheet and Refresh Chart Source Using Aspose.Cells
// Description: Shows how to build a workbook, fill the second through twentieth rows on a source sheet, attach a column chart, copy the data block to another worksheet using CopyOptions with ReferToDestinationSheet, and create a chart that automatically references the new range before saving the file.
// Keywords: Aspose.Cells | C# | .NET | CopyRows | CopyOptions | ReferToDestinationSheet | chart source range | update chart after copy | worksheet duplication | Excel automation
// Common Searches: Aspose.Cells copy rows and keep chart reference | Copy rows 2 to 20 to another sheet C# | ReferToDestinationSheet option example | Update chart data source after copying rows .NET | How to duplicate data block with chart in Aspose.Cells
// Developer Intent: Programmatically duplicate a specific row range and ensure any linked chart points to the copied data on a new worksheet.
// Use Cases: Generate a detailed report sheet that reuses a data segment and its chart without manual adjustments. | Create a template where rows are copied to a summary page and the chart automatically reflects the new range. | Automate Excel dashboards that require the same data visualisation on multiple worksheets.
// AI Prompts: Provide C# code that copies rows 2‑20 to a different worksheet with Aspose.Cells and updates the chart source automatically. | Explain the effect of the ReferToDestinationSheet flag when copying chart‑referenced rows in Aspose.Cells for .NET. | Show an example of adding a chart on a destination sheet that references a copied data range after using CopyRows with CopyOptions.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExample
{
    // Shows how to build a workbook, fill the second through twentieth rows on a source sheet, attach a column chart, copy the data block to another worksheet using CopyOptions with ReferToDestinationSheet, and create a chart that automatically references the new range before saving the file.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet as the source sheet
                Workbook workbook = new Workbook();
                Worksheet srcSheet = workbook.Worksheets[0];
                srcSheet.Name = "Source";

                // Fill rows 2‑20 (zero‑based rows 1‑19) with sample values
                for (int row = 1; row <= 19; row++)
                {
                    srcSheet.Cells[row, 0].PutValue($"Item {row}");
                    srcSheet.Cells[row, 1].PutValue(row * 10);
                }

                // Add a chart that uses the source data (A2:B20)
                int chartIdx = srcSheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
                Chart srcChart = srcSheet.Charts[chartIdx];
                srcChart.NSeries.Add("=Source!$A$2:$B$20", true);

                // -----------------------------------------------------------------
                // 1. Copy rows 2‑20 to a new worksheet
                Worksheet destSheet = workbook.Worksheets.Add("Destination");

                // Configure copy options so that any chart data source that refers to the
                // source sheet will be changed to refer to the destination sheet
                CopyOptions copyOptions = new CopyOptions
                {
                    ReferToDestinationSheet = true
                };

                // Copy 19 rows starting from row index 1 (row 2) of the source sheet
                // to row index 1 (row 2) of the destination sheet
                destSheet.Cells.CopyRows(
                    srcSheet.Cells,   // source cells
                    1,                // source start row (row 2)
                    1,                // destination start row (row 2)
                    19,               // number of rows to copy
                    copyOptions);    // apply the ReferToDestinationSheet option

                // -----------------------------------------------------------------
                // 2. Add a chart to the destination sheet that points to the copied rows
                // The data range must reference the destination sheet explicitly
                string destDataRange = $"'{destSheet.Name}'!$A$2:$B$20";
                int destChartIdx = destSheet.Charts.Add(ChartType.Column, destDataRange, true, 5, 0, 15, 5);
                Chart destChart = destSheet.Charts[destChartIdx];

                // Optionally move the chart to a different position on the sheet
                destChart.Move(10, 2, 20, 8);

                // -----------------------------------------------------------------
                // Save the workbook
                string outputPath = "RowsCopiedWithUpdatedChart.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}
