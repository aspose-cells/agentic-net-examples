// Title: Configure a PivotChart to use Tabular layout with PivotOptions in Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set pivot chart layout to tabular using PivotOptions | programmatically add pivot table and pivot chart with tabular view in .NET | show drop zones and expand collapse buttons on Aspose.Cells pivot chart | save workbook after configuring pivot chart layout with Aspose.Cells | create sample workbook and pivot chart with tabular layout using Aspose.Cells example
// Tags: pivot chart tabular layout Aspose.Cells | PivotOptions drop zones visibility | expand collapse field buttons pivot chart | create pivot table and chart C# | export workbook to xlsx Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotChartLayout
{
    // The example loads an existing workbook or creates a sample one, adds a PivotTable on a new worksheet, switches the table to Tabular form, creates a Column PivotChart linked to that table, configures the chart's PivotOptions to display all drop zones and expand/collapse buttons for a tabular view, and finally saves the workbook as an XLSX file using Aspose.Cells for .NET.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the input workbook. If the file does not exist we will create a sample workbook.
            string inputPath = "input.xlsx";
            Workbook workbook;

            if (File.Exists(inputPath))
            {
                try
                {
                    workbook = new Workbook(inputPath);
                    Console.WriteLine($"Loaded workbook from \"{inputPath}\".");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to load \"{inputPath}\": {ex.Message}");
                    Console.WriteLine("Creating a new workbook with sample data instead.");
                    workbook = CreateSampleWorkbook();
                }
            }
            else
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found. Creating a new workbook with sample data.");
                workbook = CreateSampleWorkbook();
            }

            // ------------------------------------------------------------
            // 1. Add a PivotTable based on the sample data
            // ------------------------------------------------------------
            Worksheet dataSheet = workbook.Worksheets[0];
            string dataRange = "=Sheet1!A1:C6"; // A1:C6 contains the sample data
            string pivotTableName = "SamplePivotTable";

            // Add the pivot table to a new worksheet for clarity
            Worksheet pivotSheet = workbook.Worksheets.Add("PivotSheet");
            int pivotIndex = pivotSheet.PivotTables.Add(dataRange, "A3", pivotTableName);
            PivotTable pivotTable = pivotSheet.PivotTables[pivotIndex];

            // Configure the pivot fields
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");   // Row field
            pivotTable.AddFieldToArea(PivotFieldType.Column, "Year");   // Column field
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Amount");   // Data field

            // Set the pivot table layout to Tabular form (the method exists on PivotTable)
            pivotTable.ShowInTabularForm();

            // Refresh and calculate the pivot table data
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // ------------------------------------------------------------
            // 2. Add a PivotChart that uses the newly created PivotTable
            // ------------------------------------------------------------
            int chartIndex = pivotSheet.Charts.Add(ChartType.Column, 15, 0, 30, 15);
            Chart pivotChart = pivotSheet.Charts[chartIndex];

            // Link the chart to the pivot table
            pivotChart.PivotSource = pivotTableName;

            // ------------------------------------------------------------
            // 3. Configure the PivotChart layout to Tabular using PivotOptions
            // ------------------------------------------------------------
            // The PivotOptions class controls the UI elements that appear on a pivot chart.
            // While it does not have a direct "Tabular" property, setting the drop‑zone
            // visibility together with the expand/collapse buttons gives the chart a
            // layout that matches the tabular view of the source pivot table.
            PivotOptions options = pivotChart.PivotOptions;

            // Show all pivot controls (drop zones) on the chart
            options.DropZonesVisible = true;
            options.DropZoneCategories = true;
            options.DropZoneSeries = true;
            options.DropZoneData = true;
            options.DropZoneFilter = true;

            // Show the expand/collapse buttons so the user can switch between compact
            // and tabular representations directly from the chart.
            options.ShowExpandCollapseFieldButtons = true;

            // ------------------------------------------------------------
            // 4. Save the workbook
            // ------------------------------------------------------------
            string outputPath = "output.xlsx";
            try
            {
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save workbook: {ex.Message}");
            }
        }

        private static Workbook CreateSampleWorkbook()
        {
            Workbook wb = new Workbook();
            Worksheet ws = wb.Worksheets[0];

            // Header row
            ws.Cells["A1"].PutValue("Category");
            ws.Cells["B1"].PutValue("Year");
            ws.Cells["C1"].PutValue("Amount");

            // Sample rows
            ws.Cells["A2"].PutValue("Electronics");
            ws.Cells["B2"].PutValue(2020);
            ws.Cells["C2"].PutValue(1200);

            ws.Cells["A3"].PutValue("Electronics");
            ws.Cells["B3"].PutValue(2021);
            ws.Cells["C3"].PutValue(1500);

            ws.Cells["A4"].PutValue("Furniture");
            ws.Cells["B4"].PutValue(2020);
            ws.Cells["C4"].PutValue(800);

            ws.Cells["A5"].PutValue("Furniture");
            ws.Cells["B5"].PutValue(2021);
            ws.Cells["C5"].PutValue(950);

            ws.Cells["A6"].PutValue("Office Supplies");
            ws.Cells["B6"].PutValue(2020);
            ws.Cells["C6"].PutValue(300);

            return wb;
        }
    }
}
