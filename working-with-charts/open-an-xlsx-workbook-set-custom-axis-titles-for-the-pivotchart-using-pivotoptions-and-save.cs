// Title: How to add a PivotChart with custom X‑axis and Y‑axis titles and enable drop zones using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# set X axis title on a PivotChart | Enable drop zones on a PivotChart with Aspose.Cells .NET | Create a PivotChart from a PivotTable and customize axis titles in C# | Save workbook containing a PivotChart with custom axis labels using Aspose.Cells | Add PivotChart to existing XLSX and modify chart titles programmatically
// Tags: pivot chart axis title customization Aspose.Cells | pivot chart drop zones configuration C# | create pivot chart from pivot table Aspose.Cells | export workbook with pivot chart to XLSX | chart title visibility Aspose.Cells .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Pivot;

namespace PivotChartAxisTitleDemo
{
    // The example loads or creates an XLSX workbook, builds a pivot table from sample data, adds a column PivotChart linked to that table, sets custom text and visibility for the category (X) and value (Y) axes, configures PivotOptions to display all drop zones, assigns a chart title, and saves the workbook as a new XLSX file using Aspose.Cells for .NET.
    class Program
    {
        static void Main(string[] args)
        {
            // Input and output file paths
            string inputPath = "InputWorkbook.xlsx";
            string outputPath = "PivotChartWithCustomAxisTitles.xlsx";

            try
            {
                // ------------------------------------------------------------
                // 1. Load an existing workbook or create a new one with sample data
                // ------------------------------------------------------------
                Workbook workbook;
                if (File.Exists(inputPath))
                {
                    workbook = new Workbook(inputPath);
                    Console.WriteLine($"Loaded workbook from \"{inputPath}\".");
                }
                else
                {
                    workbook = new Workbook();
                    Worksheet sheet = workbook.Worksheets[0];
                    Cells cells = sheet.Cells;

                    // Sample data for the pivot table
                    cells["A1"].PutValue("Category");
                    cells["B1"].PutValue("SubCategory");
                    cells["C1"].PutValue("Amount");

                    cells["A2"].PutValue("Fruit");
                    cells["B2"].PutValue("Apple");
                    cells["C2"].PutValue(120);

                    cells["A3"].PutValue("Fruit");
                    cells["B3"].PutValue("Banana");
                    cells["C3"].PutValue(80);

                    cells["A4"].PutValue("Vegetable");
                    cells["B4"].PutValue("Carrot");
                    cells["C4"].PutValue(50);

                    cells["A5"].PutValue("Vegetable");
                    cells["B5"].PutValue("Tomato");
                    cells["C5"].PutValue(70);

                    Console.WriteLine("Created a new workbook with sample data.");
                }

                // ------------------------------------------------------------
                // 2. Add a PivotTable based on the data range
                // ------------------------------------------------------------
                Worksheet dataSheet = workbook.Worksheets[0];
                // Define the source data range (including headers)
                string sourceData = "=Sheet1!A1:C5";

                // Add the pivot table; using the older overload (still supported)
                int pivotIndex = dataSheet.PivotTables.Add(sourceData, "E1", "SamplePivot");
                PivotTable pivotTable = dataSheet.PivotTables[pivotIndex];

                // Configure the pivot fields
                // Row: Category
                pivotTable.AddFieldToArea(PivotFieldType.Row, 0); // Column index 0 -> Category
                // Column: SubCategory
                pivotTable.AddFieldToArea(PivotFieldType.Column, 1); // Column index 1 -> SubCategory
                // Data: Amount (default aggregation is Sum for numeric fields)
                pivotTable.AddFieldToArea(PivotFieldType.Data, 2); // Column index 2 -> Amount

                // Refresh the pivot data (using the obsolete but available method)
                pivotTable.RefreshData();
                pivotTable.CalculateData();

                Console.WriteLine("PivotTable \"SamplePivot\" created and populated.");

                // ------------------------------------------------------------
                // 3. Add a PivotChart that uses the created PivotTable as its source
                // ------------------------------------------------------------
                // Add a chart somewhere else on the same sheet
                // Parameters: chart type, upper left row, upper left column, lower right row, lower right column
                int chartIndex = dataSheet.Charts.Add(ChartType.Column, 10, 0, 25, 10);
                Chart pivotChart = dataSheet.Charts[chartIndex];

                // Link the chart to the pivot table
                pivotChart.PivotSource = "SamplePivot";

                // ------------------------------------------------------------
                // 4. Set custom axis titles for the PivotChart
                // ------------------------------------------------------------
                // Category (X) axis title
                pivotChart.CategoryAxis.Title.Text = "Product Category";
                pivotChart.CategoryAxis.Title.IsVisible = true;

                // Value (Y) axis title
                pivotChart.ValueAxis.Title.Text = "Total Amount";
                pivotChart.ValueAxis.Title.IsVisible = true;

                // Optional: customize the chart title as well
                pivotChart.Title.Text = "Sales Summary by Category";
                pivotChart.Title.IsVisible = true;

                // ------------------------------------------------------------
                // 5. Demonstrate usage of PivotOptions (controls displayed on the chart)
                // ------------------------------------------------------------
                PivotOptions options = pivotChart.PivotOptions;
                options.DropZonesVisible = true;
                options.DropZoneCategories = true;
                options.DropZoneData = true;
                options.DropZoneSeries = true;
                options.DropZoneFilter = true;

                // ------------------------------------------------------------
                // 6. Save the workbook
                // ------------------------------------------------------------
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred:");
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}
