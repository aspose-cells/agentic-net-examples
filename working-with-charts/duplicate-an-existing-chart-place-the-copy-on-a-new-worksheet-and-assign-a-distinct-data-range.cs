// Title: Copy a chart to a new worksheet and bind it to a different data range with Aspose.Cells for .NET (C#)
// Description: Demonstrates how to create a workbook, add a source sheet with a column chart, then duplicate that chart on a separate worksheet, preserve its type, size and style, and assign a new data range before saving the file.
// Keywords: Aspose.Cells chart copy C# | duplicate chart Aspose.Cells | set new chart data range Aspose.Cells | clone chart properties .NET | Aspose.Cells workbook example
// Common Searches: Aspose.Cells copy chart to another sheet | C# change chart data source after duplication | how to clone a chart with Aspose.Cells | duplicate chart preserving style Aspose.Cells .NET | set chart range on copied chart Aspose.Cells
// Developer Intent: Programmatically replicate an existing chart on a different worksheet and point it to a separate data range while keeping its visual formatting.
// Use Cases: Create a quarterly report sheet that reuses the same chart layout with data for the current quarter. | Build a template workbook where each region gets its own sheet with a copy of a master chart linked to regional data. | Automate dashboard generation by copying a chart to a new tab and updating the source range for the latest metrics.
// AI Prompts: Generate C# code that copies a chart from one worksheet to another using Aspose.Cells, retains its type, dimensions, and style, and then sets a new data range. | Show how to clone a chart's visual properties and bind it to a different cell range in Aspose.Cells for .NET. | Explain step‑by‑step how to duplicate a chart on a new sheet and assign distinct source data with Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartDuplication
{
    // Demonstrates how to create a workbook, add a source sheet with a column chart, then duplicate that chart on a separate worksheet, preserve its type, size and style, and assign a new data range before saving the file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // ---------- Source worksheet with original chart ----------
            Worksheet srcSheet = workbook.Worksheets[0];
            srcSheet.Name = "Source";

            // Populate source data (A1:B4)
            srcSheet.Cells["A1"].PutValue("Category");
            srcSheet.Cells["B1"].PutValue("Value");
            srcSheet.Cells["A2"].PutValue("A");
            srcSheet.Cells["B2"].PutValue(10);
            srcSheet.Cells["A3"].PutValue("B");
            srcSheet.Cells["B3"].PutValue(20);
            srcSheet.Cells["A4"].PutValue("C");
            srcSheet.Cells["B4"].PutValue(30);

            // Add a chart to the source sheet
            int srcChartIdx = srcSheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
            Chart srcChart = srcSheet.Charts[srcChartIdx];
            srcChart.SetChartDataRange("A1:B4", true);
            srcChart.Title.Text = "Source Chart";
            srcChart.Style = 2; // Built‑in style

            // ---------- Destination worksheet ----------
            Worksheet destSheet = workbook.Worksheets.Add("CopySheet");

            // Populate distinct data for the copied chart (C1:D4)
            destSheet.Cells["C1"].PutValue("Category");
            destSheet.Cells["D1"].PutValue("Value");
            destSheet.Cells["C2"].PutValue("X");
            destSheet.Cells["D2"].PutValue(40);
            destSheet.Cells["C3"].PutValue("Y");
            destSheet.Cells["D3"].PutValue(50);
            destSheet.Cells["C4"].PutValue("Z");
            destSheet.Cells["D4"].PutValue(60);

            // Add a chart to the destination sheet with the same type and size as the source chart
            int destChartIdx = destSheet.Charts.Add(
                srcChart.Type,
                srcChart.ChartObject.UpperLeftRow,
                srcChart.ChartObject.UpperLeftColumn,
                srcChart.ChartObject.LowerRightRow,
                srcChart.ChartObject.LowerRightColumn);
            Chart destChart = destSheet.Charts[destChartIdx];

            // Assign a distinct data range to the copied chart
            destChart.SetChartDataRange("C1:D4", true);

            // Copy visual properties (title, style) from the source chart
            destChart.Title.Text = srcChart.Title.Text + " (Copy)";
            destChart.Style = srcChart.Style;

            // Save the workbook
            workbook.Save("ChartDuplicationResult.xlsx");
        }
    }
}
