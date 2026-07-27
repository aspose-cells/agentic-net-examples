using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartCloneDemo
{
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet originalSheet = workbook.Worksheets[0];
            originalSheet.Name = "Original";

            // Populate sample data for the chart
            originalSheet.Cells["A1"].PutValue("Category");
            originalSheet.Cells["A2"].PutValue("Jan");
            originalSheet.Cells["A3"].PutValue("Feb");
            originalSheet.Cells["A4"].PutValue("Mar");

            originalSheet.Cells["B1"].PutValue("Series1");
            originalSheet.Cells["B2"].PutValue(10);
            originalSheet.Cells["B3"].PutValue(20);
            originalSheet.Cells["B4"].PutValue(30);

            originalSheet.Cells["C1"].PutValue("Series2");
            originalSheet.Cells["C2"].PutValue(15);
            originalSheet.Cells["C3"].PutValue(25);
            originalSheet.Cells["C4"].PutValue(35);

            // Add a column chart to the original sheet
            int chartIndex = originalSheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart originalChart = originalSheet.Charts[chartIndex];

            // Set the data range for the chart series
            originalChart.NSeries.Add("B1:C4", true);
            originalChart.NSeries.CategoryData = "A2:A4";

            // -----------------------------------------------------------------
            // Clone the worksheet (which also clones the chart) using AddCopy
            // -----------------------------------------------------------------
            int copiedSheetIndex = workbook.Worksheets.AddCopy("Original");
            Worksheet clonedSheet = workbook.Worksheets[copiedSheetIndex];
            clonedSheet.Name = "ClonedChart";

            // Retrieve the cloned chart (same index as in the source sheet)
            Chart clonedChart = clonedSheet.Charts[chartIndex];

            // Change the series colors of the cloned chart using the ChangeColors method
            // Cast to a valid enum value; here we use 0 (Monochrome1) as an example
            clonedChart.NSeries.ChangeColors((ChartColorPaletteType)0);

            // Optionally move the cloned chart to a new location on the cloned sheet
            clonedChart.Move(10, 2, 25, 12);

            // Save the workbook to verify the result
            workbook.Save("ClonedChartDemo.xlsx");
        }
    }
}