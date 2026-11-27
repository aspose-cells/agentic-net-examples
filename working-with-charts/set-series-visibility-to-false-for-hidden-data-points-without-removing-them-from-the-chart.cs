// Title: C# Aspose.Cells: Hide rows in a column chart without removing the series using PlotVisibleCellsOnly
// Common Searches: how to exclude hidden rows from an Aspose.Cells column chart in C# | Aspose.Cells C# hide rows in chart but keep the series intact | using PlotVisibleCellsOnly to display only visible data points in an Excel chart with Aspose.Cells | C# example of a chart that ignores hidden rows in Aspose.Cells | chart series visibility control with hidden rows Aspose.Cells C#
// Tags: plotvisiblecellsonly property for Aspose.Cells chart | hide rows in Aspose.Cells column chart | preserve series while excluding hidden data points | Excel chart selective data visibility C# | Aspose.Cells chart ignore hidden rows

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsSeriesVisibilityDemo
{
    // The sample creates a workbook, populates category and value columns, hides rows that are not needed, adds a column chart covering the full range, sets PlotVisibleCellsOnly to true so hidden rows are omitted from the chart while the series remains, and saves the file as SeriesVisibilityDemo.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // -----------------------------------------------------------------
                // 1. Create a new workbook and get the first worksheet.
                // -----------------------------------------------------------------
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // -----------------------------------------------------------------
                // 2. Populate sample data.
                //    Column A – Category (used for X‑axis)
                //    Column B – Value   (used for Y‑axis)
                // -----------------------------------------------------------------
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Value");

                string[] categories = { "A", "B", "C", "D", "E" };
                int[] values = { 10, 20, 30, 40, 50 };

                for (int i = 0; i < categories.Length; i++)
                {
                    // Row indices are zero‑based; data starts at row 1 (A2, B2)
                    sheet.Cells[i + 1, 0].PutValue(categories[i]); // Column A
                    sheet.Cells[i + 1, 1].PutValue(values[i]);    // Column B

                    // Hide rows that are not "B" or "D"
                    if (categories[i] != "B" && categories[i] != "D")
                    {
                        // Use Cells.Rows collection to hide the row
                        sheet.Cells.Rows[i + 1].IsHidden = true;
                    }
                }

                // -----------------------------------------------------------------
                // 3. Add a chart that uses the whole data range.
                // -----------------------------------------------------------------
                int chartIndex = sheet.Charts.Add(ChartType.Column, 8, 0, 25, 15);
                Chart chart = sheet.Charts[chartIndex];

                // Set the data source for the series (including hidden rows)
                chart.NSeries.Add("=Sheet1!$A$2:$B$6", true);
                chart.NSeries.CategoryData = "=Sheet1!$A$2:$A$6";

                // -----------------------------------------------------------------
                // 4. Plot only visible cells.
                //    This makes the chart ignore rows that are hidden while keeping
                //    the series itself intact.
                // -----------------------------------------------------------------
                chart.PlotVisibleCellsOnly = true;

                // -----------------------------------------------------------------
                // 5. Save the workbook.
                // -----------------------------------------------------------------
                string outFile = "SeriesVisibilityDemo.xlsx";

                // Ensure the output directory exists
                string outDir = Path.GetDirectoryName(Path.GetFullPath(outFile));
                if (!Directory.Exists(outDir))
                {
                    Directory.CreateDirectory(outDir);
                }

                workbook.Save(outFile, SaveFormat.Xlsx);

                Console.WriteLine($"Workbook saved to '{outFile}'.");
                Console.WriteLine("Hidden data points are not displayed in the chart, but the series remains part of the chart.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
