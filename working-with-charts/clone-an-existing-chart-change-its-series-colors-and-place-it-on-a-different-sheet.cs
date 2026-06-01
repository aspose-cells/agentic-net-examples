using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet (source sheet)
        Workbook workbook = new Workbook();
        Worksheet sourceSheet = workbook.Worksheets[0];
        sourceSheet.Name = "Source";

        // Populate sample data for the chart
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

        // Add a column chart to the source sheet
        int chartIndex = sourceSheet.Charts.Add(ChartType.Column, 6, 0, 20, 10);
        Chart sourceChart = sourceSheet.Charts[chartIndex];
        sourceChart.NSeries.Add("B1:C4", true);          // Set data range for series
        sourceChart.NSeries.CategoryData = "A2:A4";      // Set category (X‑axis) data

        // Clone the worksheet (including the chart) to a new sheet
        int clonedSheetIndex = workbook.Worksheets.AddCopy("Source");
        Worksheet clonedSheet = workbook.Worksheets[clonedSheetIndex];
        clonedSheet.Name = "Cloned";

        // Retrieve the cloned chart (it has the same index as in the source sheet)
        Chart clonedChart = clonedSheet.Charts[chartIndex];

        // Change the series colors of the cloned chart using a monochromatic palette
        // Cast an integer to ChartColorPaletteType to avoid enum name dependencies
        clonedChart.NSeries.ChangeColors((ChartColorPaletteType)0);

        // Save the workbook with the original and cloned charts
        workbook.Save("ClonedChartDemo.xlsx");
    }
}