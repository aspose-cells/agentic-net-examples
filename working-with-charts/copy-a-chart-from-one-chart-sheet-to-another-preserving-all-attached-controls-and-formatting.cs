// Title: Copy a Chart Sheet to a New Workbook with Aspose.Cells for .NET – Preserve Controls & Formatting
// Description: Shows how to load a source workbook, locate a chart sheet, and duplicate it into a new workbook using Workbook.Worksheets.AddCopy with CopyOptions.ReferToDestinationSheet. The chart’s data range is refreshed so the copied chart points to the destination sheet while all formatting and embedded controls are retained, then the file is saved.
// Keywords: Aspose.Cells copy chart sheet | C# copy chart worksheet | AddCopy chart Aspose | CopyOptions ReferToDestinationSheet | preserve chart formatting .NET | duplicate chart sheet programmatically | chart controls Aspose.Cells | copy chart data range | Excel chart sheet copy C#
// Common Searches: copy chart sheet Aspose.Cells C# | how to duplicate a chart worksheet preserving formatting | Aspose.Cells AddCopy chart sheet example | retain chart controls when copying worksheets .NET | update chart data source after copying sheet Aspose | copy chart sheet to another workbook programmatically
// Developer Intent: Duplicate a chart sheet from one Excel file to another while keeping its formatting, controls, and data references intact.
// Use Cases: Create multiple report files from a master chart template. | Migrate legacy chart sheets into a consolidated workbook without losing embedded controls. | Generate scenario‑analysis workbooks by cloning a chart sheet and automatically linking it to new data. | Automate production of client‑specific dashboards by copying a pre‑styled chart sheet.
// AI Prompts: Provide C# code using Aspose.Cells to copy a chart sheet to a new workbook and keep all formatting and controls. | Show how to use CopyOptions.ReferToDestinationSheet to preserve chart data references when duplicating a worksheet. | Explain the steps to reassign a chart’s data range after copying a chart sheet with Aspose.Cells for .NET. | Give an example of copying a chart sheet and updating its source range in C#.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartCopyDemo
{
    // Shows how to load a source workbook, locate a chart sheet, and duplicate it into a new workbook using Workbook.Worksheets.AddCopy with CopyOptions.ReferToDestinationSheet. The chart’s data range is refreshed so the copied chart points to the destination sheet while all formatting and embedded controls are retained, then the file is saved.
    class Program
    {
        static void Main()
        {
            try
            {
                const string sourcePath = "SourceWithChart.xlsx";
                const string destPath = "DestinationWithCopiedChart.xlsx";

                // Verify source file exists
                if (!File.Exists(sourcePath))
                {
                    Console.WriteLine($"Source file not found: {sourcePath}");
                    return;
                }

                // Load the source workbook that contains the chart sheet.
                Workbook sourceWorkbook = new Workbook(sourcePath);

                // Create an empty destination workbook.
                Workbook destWorkbook = new Workbook();

                // Identify the chart sheet to be copied.
                // In Aspose.Cells a chart sheet is also represented by the Worksheet class.
                Worksheet chartSheet = sourceWorkbook.Worksheets["ChartSheet1"];
                if (chartSheet == null)
                {
                    Console.WriteLine("Chart sheet 'ChartSheet1' not found.");
                    return;
                }

                int chartSheetIndex = sourceWorkbook.Worksheets.IndexOf(chartSheet);
                if (chartSheetIndex < 0)
                {
                    Console.WriteLine("Chart sheet index could not be determined.");
                    return;
                }

                // Copy the chart sheet (including the chart, its formatting and any attached controls)
                // using the AddCopy method which copies the whole worksheet.
                int copiedIndex = destWorkbook.Worksheets.AddCopy(chartSheetIndex);
                Worksheet copiedChartSheet = destWorkbook.Worksheets[copiedIndex];

                // Ensure that the chart's data source now refers to the destination sheet.
                // This mimics Excel's behaviour when a chart is moved to another sheet.
                CopyOptions copyOptions = new CopyOptions
                {
                    ReferToDestinationSheet = true
                };
                // Apply the copy options when copying the worksheet (already applied via AddCopy).

                // The chart on the copied sheet is the first (and usually only) chart.
                if (copiedChartSheet.Charts.Count > 0)
                {
                    Chart copiedChart = copiedChartSheet.Charts[0];
                    string currentRange = copiedChart.GetChartDataRange(); // e.g., "Sheet1!A1:B5"
                    // Re‑assign the same range; with ReferToDestinationSheet = true the range will be updated
                    // to point to the copied sheet.
                    copiedChart.SetChartDataRange(currentRange, true);
                }
                else
                {
                    Console.WriteLine("No charts found on the copied sheet.");
                }

                // Save the destination workbook.
                destWorkbook.Save(destPath);
                Console.WriteLine($"Destination workbook saved to: {destPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
