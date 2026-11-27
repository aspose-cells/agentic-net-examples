// Title: C# example: Load an XLSX workbook, create a PivotTable from A1:E20, add a linked PivotChart, and save with Aspose.Cells
// Common Searches: how to generate a pivot table on sheet1 using Aspose.Cells in C# | link a column chart to a pivot table programmatically with Aspose.Cells .NET | example code for adding a pivot chart to an existing XLSX file using Aspose.Cells | C# Aspose.Cells create pivot table at G5 and chart positioned at rows 25‑45 | save an Excel workbook with pivot table and chart using Aspose.Cells for .NET
// Tags: Aspose.Cells create pivot table from range | Aspose.Cells add column PivotChart linked to pivot table | Aspose.Cells set pivot table fields row column data | Aspose.Cells apply built‑in pivot style | Aspose.Cells save workbook to XLSX format

using System;
using System.IO;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Charts;

namespace AsposeCellsPivotChartDemo
{
    // The C# program loads an existing XLSX file (or creates a sample workbook), adds a PivotTable based on the A1:E20 range positioned at G5, configures row, column, and data fields, applies a built‑in style, creates a column PivotChart linked to the PivotTable with title and legend, and saves the workbook as output.xlsx using Aspose.Cells for .NET.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            Workbook workbook = null;

            try
            {
                // Load existing workbook if it exists; otherwise create a sample workbook.
                if (File.Exists(inputPath))
                {
                    Console.WriteLine($"Loading workbook from \"{inputPath}\" ...");
                    workbook = new Workbook(inputPath);
                    Console.WriteLine("Workbook loaded successfully.");
                }
                else
                {
                    Console.WriteLine($"File \"{inputPath}\" not found. Creating a sample workbook...");
                    workbook = CreateSampleWorkbook(inputPath);
                }

                // -----------------------------------------------------------------
                // 2. Create a PivotTable from range A1:E20
                // -----------------------------------------------------------------
                Worksheet ws = workbook.Worksheets[0];
                PivotTableCollection pivots = ws.PivotTables;

                // Source data range
                string sourceRange = "=Sheet1!A1:E20";

                // Destination cell G5 (zero‑based indices: row 4, column 6)
                int destRow = 4;
                int destColumn = 6;

                Console.WriteLine("Adding PivotTable...");
                int pivotIndex = pivots.Add(sourceRange, destRow, destColumn, "MyPivotTable");
                PivotTable pivot = pivots[pivotIndex];

                // Configure fields
                pivot.AddFieldToArea(PivotFieldType.Row, 0);      // Category
                pivot.AddFieldToArea(PivotFieldType.Column, 2);   // Region
                pivot.AddFieldToArea(PivotFieldType.Data, 3);     // Sales
                pivot.DataFieldHeaderName = "Total Sales";

                // Refresh and calculate pivot data
                pivot.RefreshData();
                pivot.CalculateData();

                // Apply a built‑in style
                pivot.PivotTableStyleType = PivotTableStyleType.PivotTableStyleMedium9;
                Console.WriteLine("PivotTable created successfully.");

                // -----------------------------------------------------------------
                // 3. Add a linked PivotChart
                // -----------------------------------------------------------------
                Console.WriteLine("Adding PivotChart...");

                // Chart position (rows 25‑45, columns 0‑10)
                int chartTopRow = 25;
                int chartLeftColumn = 0;
                int chartBottomRow = 45;
                int chartRightColumn = 10;

                int chartIndex = ws.Charts.Add(ChartType.Column, chartTopRow, chartLeftColumn, chartBottomRow, chartRightColumn);
                Chart chart = ws.Charts[chartIndex];

                // Link the chart to the pivot table
                chart.PivotSource = pivot.Name;

                // Set chart title and basic formatting
                chart.Title.Text = "Sales by Region";
                chart.Legend.Position = LegendPositionType.Right;
                chart.PlotArea.Area.ForegroundColor = Color.White;

                Console.WriteLine("PivotChart added and linked to the PivotTable.");

                // -----------------------------------------------------------------
                // 4. Save the workbook
                // -----------------------------------------------------------------
                Console.WriteLine($"Saving the workbook as \"{outputPath}\" ...");
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine("Workbook saved successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
            finally
            {
                // Ensure resources are released
                workbook?.Dispose();
            }

            Console.WriteLine("Process completed. Press any key to exit.");
            Console.ReadKey();
        }

        private static Workbook CreateSampleWorkbook(string path)
        {
            var wb = new Workbook();
            Worksheet sheet = wb.Worksheets[0];
            Cells cells = sheet.Cells;

            string[] headers = { "Category", "SubCategory", "Region", "Sales", "Quantity" };
            for (int c = 0; c < headers.Length; c++)
                cells[0, c].PutValue(headers[c]);

            var rnd = new Random();
            string[] categories = { "Food", "Beverage", "Electronics" };
            string[] subCategories = { "A", "B", "C" };
            string[] regions = { "North", "South", "East", "West" };

            for (int r = 1; r <= 20; r++)
            {
                cells[r, 0].PutValue(categories[r % categories.Length]);
                cells[r, 1].PutValue(subCategories[r % subCategories.Length]);
                cells[r, 2].PutValue(regions[r % regions.Length]);
                cells[r, 3].PutValue(rnd.Next(100, 1000));   // Sales
                cells[r, 4].PutValue(rnd.Next(1, 20));      // Quantity
            }

            wb.Save(path, SaveFormat.Xlsx);
            Console.WriteLine($"Sample workbook saved as \"{path}\".");
            return wb;
        }
    }
}
