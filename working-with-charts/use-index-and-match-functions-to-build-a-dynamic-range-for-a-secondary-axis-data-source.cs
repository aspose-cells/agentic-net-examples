// Title: Create an Excel chart with a dynamic secondary‑axis series using INDEX and MATCH formulas in Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# chart secondary axis dynamic range using INDEX MATCH | How to bind a secondary axis series to a cell value with INDEX MATCH in Aspose.Cells | C# add clustered column chart with formula‑based data source in Aspose.Cells | Update secondary axis data in an Aspose.Cells chart when a worksheet cell changes | Use INDEX and MATCH to define a dynamic range for a chart series in Aspose.Cells .NET
// Tags: Aspose.Cells chart secondary axis formula | C# INDEX MATCH dynamic range Aspose.Cells | Aspose.Cells add chart with formula data source | Excel chart dynamic secondary series C# | Aspose.Cells calculate formulas for chart data

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsDynamicRangeDemo
{
    // The example creates a new workbook, writes month, sales, and profit data, and adds a clustered column chart. The primary series (Sales) uses a static range, while the secondary series (Profit) is linked to an INDEX/MATCH formula that selects the profit value for the month specified in cell E1. The series is plotted on the secondary axis, formulas are recalculated, and the workbook is saved as an XLSX file.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Path to output file (adjust as needed)
                string outputPath = "DynamicRangeSecondaryAxisChart.xlsx";

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Name = "Data";

                // Populate sample data
                string[] months = new string[]
                {
                    "Jan","Feb","Mar","Apr","May","Jun",
                    "Jul","Aug","Sep","Oct","Nov","Dec"
                };
                double[] sales = new double[] { 12000, 15000, 13000, 17000, 16000, 18000, 19000, 17500, 16500, 20000, 21000, 22000 };
                double[] profit = new double[] { 3000, 3500, 3200, 4000, 3800, 4200, 4400, 4100, 3900, 4600, 4800, 5000 };

                // Write headers
                sheet.Cells["A1"].PutValue("Month");
                sheet.Cells["B1"].PutValue("Sales");
                sheet.Cells["C1"].PutValue("Profit");

                // Fill data rows
                for (int i = 0; i < months.Length; i++)
                {
                    int row = i + 2; // Data starts at row 2
                    sheet.Cells[row, 0].PutValue(months[i]);   // Column A
                    sheet.Cells[row, 1].PutValue(sales[i]);    // Column B
                    sheet.Cells[row, 2].PutValue(profit[i]);   // Column C
                }

                // Cell E1 will hold the month we want to look up for the dynamic range
                sheet.Cells["E1"].PutValue("Jun"); // Example: user selects "Jun"

                // Add a chart (Clustered Column) to the worksheet
                int chartIndex = sheet.Charts.Add(ChartType.Column, 1, 4, 20, 12);
                Chart chart = sheet.Charts[chartIndex];
                chart.Title.Text = "Sales and Profit (Dynamic Secondary Axis)";

                // Primary axis series: Sales (static range)
                int salesSeriesIdx = chart.NSeries.Add("Data!$B$2:$B$13", true);
                Series salesSeries = chart.NSeries[salesSeriesIdx];
                salesSeries.Name = "Sales";

                // Secondary axis series: Profit (dynamic range using INDEX/MATCH)
                int profitSeriesIdx = chart.NSeries.Add(
                    "INDEX(Data!$C$2:$C$13, MATCH(Data!$E$1, Data!$A$2:$A$13, 0))",
                    true);
                Series profitSeries = chart.NSeries[profitSeriesIdx];
                profitSeries.Name = "Profit (Dynamic)";
                profitSeries.PlotOnSecondAxis = true;

                // Recalculate formulas to ensure the dynamic value is evaluated
                workbook.CalculateFormula();

                // Save the workbook
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
                Console.WriteLine("Open the file and verify that the secondary axis series updates when you change the month in cell E1.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
            }
        }
    }
}
