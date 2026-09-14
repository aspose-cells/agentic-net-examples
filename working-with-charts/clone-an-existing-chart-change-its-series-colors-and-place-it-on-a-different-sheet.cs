// Title: Clone a chart, change its series colors, and move it to another worksheet with Aspose.Cells for .NET (C#)
// AI Prompts: Using Aspose.Cells in C#, clone a worksheet that contains a chart, retrieve the cloned chart, apply a different ChartColorPaletteType to its series, and reposition the chart on the new sheet. | Generate C# code that copies a chart from one sheet, changes the series colors using a monochrome palette, and moves the chart to a specified range on a different worksheet with Aspose.Cells.
// Common Searches: Aspose.Cells C# how to copy a chart to another sheet and change its colors | change series color palette of a cloned chart in Aspose.Cells | move a duplicated chart to a new location on a different worksheet using Aspose.Cells for .NET | clone worksheet with chart and customize chart colors Aspose.Cells C# example
// Tags: clone worksheet with chart Aspose.Cells | change chart series colors ChartColorPaletteType | move chart to new range Aspose.Cells | copy chart between worksheets C# | apply monochrome palette to chart series Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a workbook, adds sample data, inserts a column chart on a source sheet, clones the worksheet (including the chart) to a new sheet, changes the cloned chart's series colors using the first ChartColorPaletteType, moves the chart to a new position on the cloned sheet, and saves the workbook as 'ClonedChartDemo.xlsx'.
class Program
{
    static void Main()
    {
        // Create a new workbook and add data for the chart
        Workbook workbook = new Workbook();
        Worksheet sourceSheet = workbook.Worksheets[0];
        sourceSheet.Name = "Source";

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

        // Add a chart to the source sheet
        int chartIndex = sourceSheet.Charts.Add(ChartType.Column, 6, 0, 20, 10);
        Chart originalChart = sourceSheet.Charts[chartIndex];
        originalChart.NSeries.Add("B1:C4", true);
        originalChart.NSeries.CategoryData = "A2:A4";

        // Clone the worksheet (which also clones the chart) to a new sheet
        int copiedIndex = workbook.Worksheets.AddCopy("Source");
        Worksheet clonedSheet = workbook.Worksheets[copiedIndex];
        clonedSheet.Name = "ClonedChartSheet";

        // Retrieve the cloned chart (same index as in the original sheet)
        Chart clonedChart = clonedSheet.Charts[chartIndex];

        // Change the series colors using a monochromatic palette
        clonedChart.NSeries.ChangeColors((ChartColorPaletteType)0); // 0 = first palette type

        // Optionally move the cloned chart to a different location on the new sheet
        clonedChart.Move(10, 2, 20, 8);

        // Save the workbook
        workbook.Save("ClonedChartDemo.xlsx");
    }
}
