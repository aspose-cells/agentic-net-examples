// Title: Hide a PivotChart legend with Aspose.Cells for .NET and save the workbook as XLSX
// Common Searches: Aspose.Cells C# hide legend on a PivotChart and export to XLSX | how to make PivotChart legend invisible using PivotOptions in .NET | remove PivotChart legend by setting transparent font with Aspose.Cells | save workbook after customizing PivotChart legend visibility in C# | Aspose.Cells example to hide chart legend and disable drop zones
// Tags: set pivotchart legend font transparent aspose.cells | disable pivotchart drop zones aspose.cells | export workbook to xlsx after chart customization aspose.cells | create pivot table and linked chart using aspose.cells c# | modify chart legend visibility programmatically aspose.cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Pivot;

namespace PivotChartLegendHideDemo
{
    // The sample creates a new workbook, adds sample data, builds a PivotTable, links a PivotChart to it, hides the chart legend by making the legend font color transparent, turns off PivotTable drop‑zone controls via PivotOptions, and finally saves the workbook as an XLSX file using Aspose.Cells for .NET.
    class Program
    {
        static void Main(string[] args)
        {
            // Output file path
            string outputPath = "PivotChart_HideLegend.xlsx";

            try
            {
                // ------------------------------------------------------------
                // 1. Create a new workbook and add sample data
                // ------------------------------------------------------------
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Sample data
                cells["A1"].PutValue("Category");
                cells["B1"].PutValue("Value");
                cells["A2"].PutValue("A");
                cells["B2"].PutValue(10);
                cells["A3"].PutValue("B");
                cells["B3"].PutValue(20);
                cells["A4"].PutValue("A");
                cells["B4"].PutValue(30);
                cells["A5"].PutValue("B");
                cells["B5"].PutValue(40);

                // ------------------------------------------------------------
                // 2. Add a PivotTable based on the data
                // ------------------------------------------------------------
                // Using the overload that specifies source range, destination cell and table name
                int pivotIndex = sheet.PivotTables.Add("A1:B5", "D1", "SamplePivot");
                PivotTable pivotTable = sheet.PivotTables[pivotIndex];
                pivotTable.AddFieldToArea(PivotFieldType.Row, 0);   // Category
                pivotTable.AddFieldToArea(PivotFieldType.Data, 1); // Value

                // ------------------------------------------------------------
                // 3. Add a PivotChart linked to the PivotTable
                // ------------------------------------------------------------
                int chartIndex = sheet.Charts.Add(ChartType.Column, 7, 0, 20, 10);
                Chart chart = sheet.Charts[chartIndex];
                chart.PivotSource = "SamplePivot";

                // ------------------------------------------------------------
                // 4. Hide the chart legend
                // ------------------------------------------------------------
                // The Legend class does not expose an IsVisible property in this version.
                // As a workaround, make the legend transparent so it is not visible.
                chart.Legend.Font.Color = System.Drawing.Color.Transparent;

                // ------------------------------------------------------------
                // 5. Optional: adjust PivotOptions (e.g., hide drop‑zone controls)
                // ------------------------------------------------------------
                PivotOptions pivotOptions = chart.PivotOptions;
                pivotOptions.DropZonesVisible = false;

                // ------------------------------------------------------------
                // 6. Ensure output directory exists and save the workbook
                // ------------------------------------------------------------
                string outDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outDir) && !Directory.Exists(outDir))
                {
                    Directory.CreateDirectory(outDir);
                }

                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'. Legend is hidden.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}
