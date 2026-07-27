using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartDuplication
{
    class Program
    {
        static void Main()
        {
            // ------------------------------------------------------------
            // 1. Create a source workbook with sample data and a chart
            // ------------------------------------------------------------
            Workbook srcWorkbook = new Workbook();
            Worksheet srcSheet = srcWorkbook.Worksheets[0];
            srcSheet.Name = "SourceSheet";

            // Populate source data (A1:B4)
            srcSheet.Cells["A1"].PutValue("Category");
            srcSheet.Cells["B1"].PutValue("Value");
            srcSheet.Cells["A2"].PutValue("A");
            srcSheet.Cells["B2"].PutValue(10);
            srcSheet.Cells["A3"].PutValue("B");
            srcSheet.Cells["B3"].PutValue(20);
            srcSheet.Cells["A4"].PutValue("C");
            srcSheet.Cells["B4"].PutValue(30);

            // Add a chart that uses the above data
            int chartIdx = srcSheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
            Chart srcChart = srcSheet.Charts[chartIdx];
            srcChart.SetChartDataRange("A1:B4", true);
            srcChart.Title.Text = "Source Chart";

            // ------------------------------------------------------------
            // 2. Add a new worksheet that will receive the duplicated chart
            // ------------------------------------------------------------
            Worksheet destSheet = srcWorkbook.Worksheets.Add("CopiedChartSheet");

            // ------------------------------------------------------------
            // 3. Copy the source worksheet (including the chart) to the new sheet
            //    Use CopyOptions.ReferToDestinationSheet = true so that the
            //    chart's data source points to the destination sheet.
            // ------------------------------------------------------------
            CopyOptions copyOpts = new CopyOptions();
            copyOpts.ReferToDestinationSheet = true;   // chart will refer to destSheet
            srcSheet.Copy(destSheet, copyOpts);

            // ------------------------------------------------------------
            // 4. Prepare a distinct data range on the destination sheet
            //    (e.g., cells A10:B13) and fill it with new values.
            // ------------------------------------------------------------
            destSheet.Cells["A10"].PutValue("Category");
            destSheet.Cells["B10"].PutValue("Value");
            destSheet.Cells["A11"].PutValue("X");
            destSheet.Cells["B11"].PutValue(40);
            destSheet.Cells["A12"].PutValue("Y");
            destSheet.Cells["B12"].PutValue(50);
            destSheet.Cells["A13"].PutValue("Z");
            destSheet.Cells["B13"].PutValue(60);

            // ------------------------------------------------------------
            // 5. Locate the copied chart in the destination sheet and assign
            //    the new data range to it.
            // ------------------------------------------------------------
            if (destSheet.Charts.Count > 0)
            {
                Chart copiedChart = destSheet.Charts[0];
                copiedChart.SetChartDataRange("A10:B13", true);
                copiedChart.Title.Text = "Copied Chart with New Data";
            }

            // ------------------------------------------------------------
            // 6. Save the workbook
            // ------------------------------------------------------------
            srcWorkbook.Save("ChartDuplicationResult.xlsx");
        }
    }
}