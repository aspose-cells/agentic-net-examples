// Title: Cut a range that contains a chart and paste it into a new workbook with Aspose.Cells for .NET
// Description: Demonstrates how to copy a cell range that includes a chart, preserve the chart and its source data, and then delete the original range to achieve a cut operation across workbooks using Aspose.Cells for .NET.
// Keywords: Aspose.Cells cut range | copy chart with data | PasteOptions All | KeepOldTables | move chart between workbooks C# | .NET spreadsheet chart transfer | delete range after copy
// Common Searches: cut range with chart Aspose.Cells | copy chart and data to another workbook .NET | preserve chart when moving range Aspose.Cells | PasteOptions KeepOldTables example | remove original range after copying chart
// Developer Intent: Move a chart and its underlying data from one workbook to another while keeping the chart functional and removing the original block.
// Use Cases: Extract a chart section from a master report and place it in a summary workbook. | Generate a report by copying a templated chart area into a new file and clearing the template. | Automate consolidation of charted data blocks across multiple spreadsheets.
// AI Prompts: Write C# code with Aspose.Cells that cuts a range containing a chart and pastes it into a new workbook, preserving chart links. | Show how to use PasteOptions (PasteType.All, KeepOldTables) to copy a chart range and then delete the source range. | Explain how chart references are updated automatically when a range with a chart is moved to another workbook using Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using AsposeRange = Aspose.Cells.Range;

// Demonstrates how to copy a cell range that includes a chart, preserve the chart and its source data, and then delete the original range to achieve a cut operation across workbooks using Aspose.Cells for .NET.
class CutRangeWithChart
{
    static void Main()
    {
        try
        {
            // Create source workbook and populate data
            Workbook srcWb = new Workbook();
            Worksheet srcSheet = srcWb.Worksheets[0];
            srcSheet.Name = "Source";

            srcSheet.Cells["A1"].PutValue("Category");
            srcSheet.Cells["B1"].PutValue("Value");
            for (int i = 2; i <= 6; i++)
            {
                srcSheet.Cells[$"A{i}"].PutValue($"Item {i - 1}");
                srcSheet.Cells[$"B{i}"].PutValue(i * 10);
            }

            // Add a column chart referencing the data
            int chartIndex = srcSheet.Charts.Add(ChartType.Column, 7, 0, 20, 5);
            Chart chart = srcSheet.Charts[chartIndex];
            chart.NSeries.Add("=Source!$B$2:$B$6", true);
            chart.NSeries.CategoryData = "=Source!$A$2:$A$6";

            // Define the range that includes data and chart (rows 0‑20, cols 0‑5)
            int startRow = 0;
            int startCol = 0;
            int rowCount = 21;
            int colCount = 6;
            AsposeRange srcRange = srcSheet.Cells.CreateRange(startRow, startCol, rowCount, colCount);

            // Destination workbook
            Workbook destWb = new Workbook();
            Worksheet destSheet = destWb.Worksheets[0];
            destSheet.Name = "Destination";

            AsposeRange destRange = destSheet.Cells.CreateRange(startRow, startCol, rowCount, colCount);

            // Copy (cut) the range
            PasteOptions pasteOptions = new PasteOptions
            {
                PasteType = PasteType.All,
                KeepOldTables = true
            };
            destRange.Copy(srcRange, pasteOptions);

            // Delete original range to achieve cut effect
            srcSheet.Cells.DeleteRange(startRow, startCol, rowCount, colCount, ShiftType.Down);

            // Save workbooks
            srcWb.Save("SourceAfterCut.xlsx");
            destWb.Save("DestinationWithChart.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
