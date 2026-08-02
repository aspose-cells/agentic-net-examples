using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    public class LinkChartSeriesToNamedRange
    {
        public static void Run()
        {
            try
            {
                // 1. Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // 2. Populate sample data that will be used for the chart
                //    Column A – Category (e.g., months)
                //    Column B – Sales values
                sheet.Cells["A1"].PutValue("Month");
                sheet.Cells["B1"].PutValue("Sales");
                string[] months = { "Jan", "Feb", "Mar", "Apr", "May" };
                int[] sales = { 120, 150, 130, 170, 160 };

                for (int i = 0; i < months.Length; i++)
                {
                    sheet.Cells[i + 2, 0].PutValue(months[i]);   // A column
                    sheet.Cells[i + 2, 1].PutValue(sales[i]);   // B column
                }

                // 3. Define a named range called "SalesData" that refers to the sales values (B2:B6)
                int nameIndex = workbook.Worksheets.Names.Add("SalesData");
                workbook.Worksheets.Names[nameIndex].RefersTo = $"={sheet.Name}!$B$2:$B$6";

                // 4. Add a chart to the worksheet
                int chartIndex = sheet.Charts.Add(ChartType.Column, 7, 0, 20, 10);
                Chart chart = sheet.Charts[chartIndex];

                // 5. Link the chart series to the named range "SalesData"
                chart.NSeries.Add("SalesData", true);

                // 6. (Optional) Set the category (X‑axis) data to the month names
                chart.NSeries.CategoryData = $"={sheet.Name}!$A$2:$A$6";

                // 7. Give the series a readable name (linked to the header cell)
                chart.NSeries[0].Name = $"={sheet.Name}!$B$1";

                // 8. Save the workbook
                string outputPath = "ChartWithNamedRange.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            LinkChartSeriesToNamedRange.Run();
        }
    }
}