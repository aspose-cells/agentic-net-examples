// Title: Clone a Chart, Change Its Series Colors, and Move It to Another Worksheet – Aspose.Cells for .NET
// Description: This Aspose.Cells for .NET example shows how to create a workbook, add a column chart, duplicate the worksheet (which copies the chart), modify the cloned chart's series colors with Chart.NSeries.ChangeColors, and reposition the chart on the new sheet before saving the file.
// Keywords: Aspose.Cells clone chart | change chart series colors | move chart to another worksheet | copy worksheet with charts | ChartColorPaletteType | C# Aspose.Cells chart example | duplicate chart Aspose.Cells
// Common Searches: how to duplicate a chart in Aspose.Cells .NET | change series colors of a cloned chart Aspose.Cells | move chart to a different sheet using Aspose.Cells | copy worksheet containing charts Aspose.Cells | Aspose.Cells Chart.NSeries.ChangeColors usage
// Developer Intent: Copy an existing chart, recolor its series, and place it on a different worksheet.
// Use Cases: Reuse a standard chart template across multiple report tabs while applying a unique color scheme per region. | Generate a consolidated workbook where each sheet shows a cloned chart with brand‑specific colors. | Build a dashboard that copies source charts, updates their palettes to match a theme, and arranges them on a summary sheet.
// AI Prompts: Write C# code with Aspose.Cells to clone a chart from one worksheet, apply a custom ChartColorPaletteType, and move the chart to another worksheet. | Explain the Chart.NSeries.ChangeColors method, list available ChartColorPaletteType values, and demonstrate selecting a specific palette. | Show how to copy a worksheet preserving all embedded charts, rename the cloned sheet, and reposition the cloned chart using Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartCloneDemo
{
    // This Aspose.Cells for .NET example shows how to create a workbook, add a column chart, duplicate the worksheet (which copies the chart), modify the cloned chart's series colors with Chart.NSeries.ChangeColors, and reposition the chart on the new sheet before saving the file.
    class Program
    {
        static void Main()
        {
            // ---------- Create a workbook and populate data ----------
            Workbook workbook = new Workbook();
            Worksheet sourceSheet = workbook.Worksheets[0];
            sourceSheet.Name = "SourceSheet";

            // Sample data for the chart
            sourceSheet.Cells["A1"].PutValue("Category");
            sourceSheet.Cells["A2"].PutValue("Jan");
            sourceSheet.Cells["A3"].PutValue("Feb");
            sourceSheet.Cells["A4"].PutValue("Mar");

            sourceSheet.Cells["B1"].PutValue("Series1");
            sourceSheet.Cells["B2"].PutValue(10);
            sourceSheet.Cells["B3"].PutValue(20);
            sourceSheet.Cells["B4"].PutValue(30);

            sourceSheet.Cells["C1"].PutValue("Series2");
            sourceSheet.Cells["C2"].PutValue(15);
            sourceSheet.Cells["C3"].PutValue(25);
            sourceSheet.Cells["C4"].PutValue(35);

            // Add an original chart on the source sheet
            int originalChartIdx = sourceSheet.Charts.Add(ChartType.Column, 6, 0, 20, 10);
            Chart originalChart = sourceSheet.Charts[originalChartIdx];
            originalChart.NSeries.Add("B1:C4", true);          // data series
            originalChart.NSeries.CategoryData = "A2:A4";     // categories

            // ---------- Clone the worksheet (which also clones the chart) ----------
            // Use the AddCopy(string) rule to copy the source worksheet.
            int copiedSheetIdx = workbook.Worksheets.AddCopy("SourceSheet");
            Worksheet copiedSheet = workbook.Worksheets[copiedSheetIdx];
            copiedSheet.Name = "ClonedSheet";

            // Retrieve the cloned chart (same index as original because the copy preserves order)
            Chart clonedChart = copiedSheet.Charts[originalChartIdx];

            // ---------- Change series colors of the cloned chart ----------
            // Use SeriesCollection.ChangeColors method. Cast to a valid enum value (e.g., 0).
            clonedChart.NSeries.ChangeColors((ChartColorPaletteType)0);

            // Optional: move the cloned chart to a different location on the new sheet
            clonedChart.Move(10, 2, 20, 8); // topRow, leftColumn, bottomRow, rightColumn

            // ---------- Save the workbook ----------
            workbook.Save("ClonedChartDemo.xlsx");
        }
    }
}
