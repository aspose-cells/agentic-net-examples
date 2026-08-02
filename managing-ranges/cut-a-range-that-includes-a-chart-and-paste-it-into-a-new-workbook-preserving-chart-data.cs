// Title: Cut a range that includes a chart and paste it into a new workbook with Aspose.Cells for .NET
// Description: Demonstrates how to copy a cell range that contains a chart, paste it into another workbook while preserving the chart and its data source, and then remove the original range and chart to emulate a cut operation. The example uses Aspose.Cells' PasteOptions with PasteType.All and shows saving both the source and destination workbooks.
// Keywords: Aspose.Cells copy range with chart | cut and paste chart between workbooks .NET | preserve chart data Aspose.Cells | PasteOptions PasteType.All example | copy drawings and charts Aspose.Cells | C# Aspose.Cells chart migration | move chart and data to new workbook
// Common Searches: Aspose.Cells copy range that includes a chart | cut range with chart and paste to another workbook .NET | preserve chart objects when moving cells Aspose.Cells | how to use PasteOptions to copy charts in C# | remove original chart after copying range Aspose.Cells
// Developer Intent: Copy a range containing a chart to a new workbook and delete the original range and chart to achieve a cut‑and‑paste effect.
// Use Cases: Extract a chart and its source data from a template to generate a standalone report file. | Automate the relocation of a chart section to a separate workbook for distribution while cleaning up the source file. | Create identical chart areas across multiple workbooks to ensure visual consistency in automated reporting.
// AI Prompts: Write C# code using Aspose.Cells to cut a range that includes a chart and paste it into a new workbook, keeping the chart linked to its data. | Show how to configure PasteOptions with PasteType.All to copy cells, formats, drawings, and charts between workbooks in Aspose.Cells. | Explain the steps to delete the original chart and range after copying them to simulate a cut operation in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using AsposeRange = Aspose.Cells.Range;

// Demonstrates how to copy a cell range that contains a chart, paste it into another workbook while preserving the chart and its data source, and then remove the original range and chart to emulate a cut operation. The example uses Aspose.Cells' PasteOptions with PasteType.All and shows saving both the source and destination workbooks.
class Program
{
    static void Main()
    {
        try
        {
            // ---------- Create source workbook with data and a chart ----------
            Workbook srcWb = new Workbook();
            Worksheet srcSheet = srcWb.Worksheets[0];

            // Populate sample data
            srcSheet.Cells["A1"].PutValue("Category");
            srcSheet.Cells["B1"].PutValue("Value");
            for (int i = 0; i < 5; i++)
            {
                srcSheet.Cells[i + 1, 0].PutValue($"Item {i + 1}");
                srcSheet.Cells[i + 1, 1].PutValue((i + 1) * 10);
            }

            // Add a column chart anchored to rows 0‑10 and columns 2‑7
            int chartIdx = srcSheet.Charts.Add(ChartType.Column, 0, 2, 10, 7);
            Chart chart = srcSheet.Charts[chartIdx];
            chart.NSeries.Add("B2:B6", true);               // Values
            chart.NSeries.CategoryData = "A2:A6";           // Categories

            // Define source range that includes both data and chart area (11 rows × 8 columns)
            AsposeRange srcRange = srcSheet.Cells.CreateRange(0, 0, 11, 8);

            // ---------- Create destination workbook ----------
            Workbook destWb = new Workbook();
            Worksheet destSheet = destWb.Worksheets[0];
            AsposeRange destRange = destSheet.Cells.CreateRange(0, 0, 11, 8);

            // Paste options – copy everything (values, formats, drawings, charts, etc.)
            PasteOptions pasteOptions = new PasteOptions
            {
                PasteType = PasteType.All
            };

            // ---------- Cut (copy then remove) ----------
            // Copy the range with the chart to the destination workbook
            destRange.Copy(srcRange, pasteOptions);

            // Remove the original range from the source sheet to simulate a cut operation
            srcSheet.Cells.DeleteRange(0, 0, 11, 8, ShiftType.None);

            // Remove the original chart object (optional, because range deletion does not delete drawings)
            srcSheet.Charts.RemoveAt(chartIdx);

            // ---------- Save both workbooks ----------
            srcWb.Save("SourceAfterCut.xlsx");
            destWb.Save("DestinationWithChart.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
