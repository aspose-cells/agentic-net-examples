using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsDashboardExample
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook (lifecycle: create)
                Workbook workbook = new Workbook();

                // -------------------------------------------------
                // Prepare sample data on two source worksheets
                // -------------------------------------------------
                // Sheet1 (default worksheet)
                Worksheet sheet1 = workbook.Worksheets[0];
                sheet1.Name = "RegionA";
                sheet1.Cells["A1"].PutValue("Product");
                sheet1.Cells["B1"].PutValue("Sales");
                sheet1.Cells["A2"].PutValue("Apple");
                sheet1.Cells["B2"].PutValue(120);
                sheet1.Cells["A3"].PutValue("Banana");
                sheet1.Cells["B3"].PutValue(80);
                sheet1.Cells["A4"].PutValue("Cherry");
                sheet1.Cells["B4"].PutValue(150);

                // Sheet2
                int sheet2Index = workbook.Worksheets.Add();
                Worksheet sheet2 = workbook.Worksheets[sheet2Index];
                sheet2.Name = "RegionB";
                sheet2.Cells["A1"].PutValue("Product");
                sheet2.Cells["B1"].PutValue("Sales");
                sheet2.Cells["A2"].PutValue("Apple");
                sheet2.Cells["B2"].PutValue(200);
                sheet2.Cells["A3"].PutValue("Banana");
                sheet2.Cells["B3"].PutValue(130);
                sheet2.Cells["A4"].PutValue("Cherry");
                sheet2.Cells["B4"].PutValue(170);

                // -------------------------------------------------
                // Create a named range that aggregates sales from both sheets
                // The name refers to a SUM formula that adds the same range on each sheet
                // -------------------------------------------------
                int nameIndex = workbook.Worksheets.Names.Add("TotalSales");
                Name totalSalesName = workbook.Worksheets.Names[nameIndex];
                // RefersTo must start with '=' and can contain multiple sheet references
                totalSalesName.RefersTo = "=SUM(RegionA!$B$2:$B$4,RegionB!$B$2:$B$4)";

                // -------------------------------------------------
                // Use the named range on a dashboard sheet
                // -------------------------------------------------
                int dashboardIndex = workbook.Worksheets.Add();
                Worksheet dashboard = workbook.Worksheets[dashboardIndex];
                dashboard.Name = "Dashboard";
                dashboard.Cells["A1"].PutValue("Aggregated Sales (All Regions):");
                // The cell formula simply references the named range
                dashboard.Cells["B1"].Formula = "=TotalSales";

                // Calculate all formulas so the dashboard shows the result
                workbook.CalculateFormula();

                // -------------------------------------------------
                // Save the workbook (lifecycle: save)
                // -------------------------------------------------
                string outputPath = "DashboardAggregatedSales.xlsx";

                // Ensure the directory exists before saving
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                // Log the exception details for troubleshooting
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}