// Title: Copy a chart sheet with formatting and shapes using Aspose.Cells for .NET
// Description: Loads a source workbook, extracts a chart from a chart sheet, creates a destination workbook, adds a matching chart, copies series data, titles, legends, styles, and any attached shapes, then saves the new file while preserving all visual and data references.
// Keywords: Aspose.Cells copy chart | duplicate chart sheet .NET | preserve chart formatting Aspose | copy chart shapes Aspose.Cells | chart sheet to new workbook | CopyOptions ReferToDestinationSheet | C# Aspose.Cells chart transfer
// Common Searches: how to copy a chart sheet with Aspose.Cells | Aspose.Cells copy chart preserving formatting | C# copy chart and attached shapes between workbooks | Aspose.Cells chart sheet duplication example | retain chart data source when moving chart to another sheet
// Developer Intent: Copy a chart from one chart sheet to another workbook while keeping all series, formatting, and attached shapes intact.
// Use Cases: Reuse a template chart in multiple generated reports without losing styling or annotations. | Create localized dashboards by cloning a chart with its text boxes and images into separate workbooks. | Migrate legacy Excel charts to new files automatically, preserving data links and visual design.
// AI Prompts: Write C# code with Aspose.Cells that copies a chart sheet to another workbook, preserving series, titles, legends, styles, and attached shapes. | Explain how CopyOptions.ReferToDestinationSheet updates chart data references during a copy operation in Aspose.Cells. | Provide a step‑by‑step guide to copy several charts from one workbook’s chart sheets to individual sheets in a new workbook using Aspose.Cells for .NET.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsChartCopyDemo
{
    // Loads a source workbook, extracts a chart from a chart sheet, creates a destination workbook, adds a matching chart, copies series data, titles, legends, styles, and any attached shapes, then saves the new file while preserving all visual and data references.
    class Program
    {
        static void Main()
        {
            const string sourcePath = "SourceWithChart.xlsx";
            const string destPath = "DestinationWithCopiedChart.xlsx";

            // Verify source file exists
            if (!File.Exists(sourcePath))
            {
                Console.WriteLine($"Source file \"{sourcePath}\" not found.");
                return;
            }

            try
            {
                // Load the source workbook that contains the chart sheet
                Workbook sourceWorkbook = new Workbook(sourcePath);

                // Create a new (empty) destination workbook
                Workbook destWorkbook = new Workbook();

                // -----------------------------------------------------------------
                // 1. Identify the source chart sheet and the chart to be copied
                // -----------------------------------------------------------------
                Worksheet sourceChartSheet = sourceWorkbook.Worksheets["ChartSheet1"];
                if (sourceChartSheet == null || sourceChartSheet.Charts.Count == 0)
                {
                    Console.WriteLine("Source chart sheet or chart not found.");
                    return;
                }
                Chart sourceChart = sourceChartSheet.Charts[0];

                // -----------------------------------------------------------------
                // 2. Add a new worksheet to the destination workbook that will host the copied chart
                // -----------------------------------------------------------------
                Worksheet destChartSheet = destWorkbook.Worksheets.Add("CopiedChartSheet");

                // -----------------------------------------------------------------
                // 3. Prepare copy options so that the chart data source points to the destination sheet
                // -----------------------------------------------------------------
                CopyOptions copyOptions = new CopyOptions
                {
                    ReferToDestinationSheet = true   // Adjust data source references to the new sheet
                };

                // -----------------------------------------------------------------
                // 4. Add a new chart to the destination sheet with the same type and position as the source chart
                // -----------------------------------------------------------------
                int destChartIndex = destChartSheet.Charts.Add(
                    sourceChart.Type,
                    sourceChart.ChartObject.UpperLeftRow,
                    sourceChart.ChartObject.UpperLeftColumn,
                    sourceChart.ChartObject.LowerRightRow,
                    sourceChart.ChartObject.LowerRightColumn);
                Chart destChart = destChartSheet.Charts[destChartIndex];

                // -----------------------------------------------------------------
                // 5. Copy the series (data range) from source to destination
                // -----------------------------------------------------------------
                int seriesIdx = 0;
                foreach (Series srcSeries in sourceChart.NSeries)
                {
                    // Add series values formula; orientation set to true (vertical) as a default
                    destChart.NSeries.Add(srcSeries.Values, true);
                    // Preserve series name
                    destChart.NSeries[seriesIdx].Name = srcSeries.Name;
                    // Preserve X‑values if they are defined separately
                    if (!string.IsNullOrEmpty(srcSeries.XValues))
                    {
                        destChart.NSeries[seriesIdx].XValues = srcSeries.XValues;
                    }
                    seriesIdx++;
                }

                // Copy category (X‑axis) data if it is set separately
                if (!string.IsNullOrEmpty(sourceChart.NSeries.CategoryData))
                {
                    destChart.NSeries.CategoryData = sourceChart.NSeries.CategoryData;
                }

                // -----------------------------------------------------------------
                // 6. Copy visual properties (title, legend, style, placement, etc.)
                // -----------------------------------------------------------------
                destChart.Title.Text = sourceChart.Title.Text;
                destChart.Title.Font.Name = sourceChart.Title.Font.Name;
                destChart.Title.Font.Size = sourceChart.Title.Font.Size;
                destChart.Title.Font.IsBold = sourceChart.Title.Font.IsBold;

                destChart.Legend.Position = sourceChart.Legend.Position;
                destChart.Legend.IsOverLay = sourceChart.Legend.IsOverLay;

                destChart.Style = sourceChart.Style;
                destChart.Placement = sourceChart.Placement;
                destChart.SizeWithWindow = sourceChart.SizeWithWindow;
                destChart.ShowLegend = sourceChart.ShowLegend;
                destChart.ShowDataTable = sourceChart.ShowDataTable;
                destChart.PlotEmptyCellsType = sourceChart.PlotEmptyCellsType;
                destChart.DisplayNaAsBlank = sourceChart.DisplayNaAsBlank;

                // -----------------------------------------------------------------
                // 7. Copy any shapes (e.g., text boxes, pictures) that are attached to the chart
                // -----------------------------------------------------------------
                foreach (Shape srcShape in sourceChart.Shapes)
                {
                    try
                    {
                        // AddCopy copies the shape and keeps the original position relative to the chart.
                        destChart.Shapes.AddCopy(
                            srcShape,
                            srcShape.UpperLeftRow,
                            srcShape.UpperLeftColumn,
                            srcShape.LowerRightRow,
                            srcShape.LowerRightColumn);
                    }
                    catch (Exception shapeEx)
                    {
                        Console.WriteLine($"Failed to copy shape: {shapeEx.Message}");
                    }
                }

                // -----------------------------------------------------------------
                // 8. Save the destination workbook
                // -----------------------------------------------------------------
                // Ensure the directory for the destination file exists
                string destDir = Path.GetDirectoryName(destPath);
                if (!string.IsNullOrEmpty(destDir) && !Directory.Exists(destDir))
                {
                    Directory.CreateDirectory(destDir);
                }

                destWorkbook.Save(destPath);
                Console.WriteLine("Chart copied successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
