// Title: Load an XLS workbook, add a PivotChart with percentage data labels via PivotOptions, and save using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# show percentages on a PivotChart in an existing XLS file | enable drop zones for a PivotChart using Aspose.Cells .NET | create or retrieve a PivotTable and link it to a chart with Aspose.Cells | save modified XLS workbook with refreshed chart cache using Aspose.Cells | set data labels to display only percentages on a pie chart in Aspose.Cells
// Tags: Aspose.Cells load XLS workbook | create pivot table Aspose.Cells C# | add pivot chart linked to pivot table Aspose.Cells | configure PivotOptions drop zones Aspose.Cells | show percentage data labels Aspose.Cells chart | XlsSaveOptions refresh chart cache

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Pivot;

namespace PivotChartPercentageDemo
{
    // The example loads an existing XLS workbook, creates a PivotTable if none is present, adds a pie PivotChart linked to that table, enables all PivotOptions drop zones, configures the series to display only percentage data labels, and saves the workbook back to XLS using XlsSaveOptions with chart cache refresh.
    class Program
    {
        static void Main(string[] args)
        {
            // Define input and output file paths
            string inputPath = "input.xls";
            string outputPath = "output.xls";

            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" does not exist.");
                return;
            }

            try
            {
                // Load the existing XLS workbook
                Workbook workbook = new Workbook(inputPath);
                Worksheet sheet = workbook.Worksheets[0];

                // ------------------------------------------------------------
                // 1. Ensure there is a PivotTable in the worksheet.
                //    If none exists, create a simple one for demonstration.
                // ------------------------------------------------------------
                PivotTable pivotTable = null;
                if (sheet.PivotTables.Count > 0)
                {
                    pivotTable = sheet.PivotTables[0];
                }
                else
                {
                    // Create sample data for a pivot table
                    sheet.Cells["A1"].PutValue("Category");
                    sheet.Cells["B1"].PutValue("SubCategory");
                    sheet.Cells["C1"].PutValue("Amount");

                    sheet.Cells["A2"].PutValue("Fruit");
                    sheet.Cells["B2"].PutValue("Apple");
                    sheet.Cells["C2"].PutValue(120);

                    sheet.Cells["A3"].PutValue("Fruit");
                    sheet.Cells["B3"].PutValue("Banana");
                    sheet.Cells["C3"].PutValue(80);

                    sheet.Cells["A4"].PutValue("Vegetable");
                    sheet.Cells["B4"].PutValue("Carrot");
                    sheet.Cells["C4"].PutValue(50);

                    sheet.Cells["A5"].PutValue("Vegetable");
                    sheet.Cells["B5"].PutValue("Broccoli");
                    sheet.Cells["C5"].PutValue(70);

                    // Add a pivot table based on the sample data
                    // Note: This overload is obsolete but still functional.
                    int pivotIndex = sheet.PivotTables.Add("=A1:C5", "E1", "DemoPivot");
                    pivotTable = sheet.PivotTables[pivotIndex];
                    pivotTable.AddFieldToArea(PivotFieldType.Row, 0);   // Category
                    pivotTable.AddFieldToArea(PivotFieldType.Row, 1);   // SubCategory
                    pivotTable.AddFieldToArea(PivotFieldType.Data, 2);  // Amount
                    pivotTable.DataFieldHeaderName = "Total Amount";

                    // Refresh and calculate data
                    pivotTable.RefreshData();
                    pivotTable.CalculateData();
                }

                if (pivotTable == null)
                {
                    throw new InvalidOperationException("Failed to create or retrieve a PivotTable.");
                }

                // ------------------------------------------------------------
                // 2. Add a PivotChart that uses the created PivotTable as its source.
                // ------------------------------------------------------------
                int chartIndex = sheet.Charts.Add(ChartType.Pie, 10, 0, 25, 15);
                Chart chart = sheet.Charts[chartIndex];
                chart.PivotSource = pivotTable.Name; // Link chart to the pivot table

                // ------------------------------------------------------------
                // 3. Configure PivotOptions (required by the task) – enable drop zones.
                // ------------------------------------------------------------
                PivotOptions pivotOptions = chart.PivotOptions;
                pivotOptions.DropZonesVisible = true;
                pivotOptions.DropZoneCategories = true;
                pivotOptions.DropZoneData = true;
                pivotOptions.DropZoneSeries = true;
                pivotOptions.DropZoneFilter = true;

                // ------------------------------------------------------------
                // 4. Show data as percentages.
                //    For a pie chart the percentages are shown automatically,
                //    but for other chart types we explicitly enable them.
                // ------------------------------------------------------------
                foreach (Series series in chart.NSeries)
                {
                    series.DataLabels.ShowPercentage = true;
                    series.DataLabels.ShowValue = false;
                }

                // ------------------------------------------------------------
                // 5. Save the workbook using XlsSaveOptions to keep the original format.
                // ------------------------------------------------------------
                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                XlsSaveOptions saveOptions = new XlsSaveOptions
                {
                    RefreshChartCache = true
                };

                workbook.Save(outputPath, saveOptions);
                Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred while processing the workbook:");
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}
