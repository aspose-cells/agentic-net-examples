using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // ----- Populate sample sales data -----
            // Header
            sheet.Cells["A1"].PutValue("Region");
            sheet.Cells["B1"].PutValue("Product A");
            sheet.Cells["C1"].PutValue("Product B");

            // Data (Product B values will be stored as negative to create the tornado effect)
            string[] regions = { "North", "South", "East", "West", "Central" };
            int[] productA = { 120, 80, 150, 70, 110 };
            int[] productB = { 100, 90, 130, 60, 95 };

            for (int i = 0; i < regions.Length; i++)
            {
                sheet.Cells[i + 1, 0].PutValue(regions[i]);          // Region name
                sheet.Cells[i + 1, 1].PutValue(productA[i]);        // Product A (positive)
                sheet.Cells[i + 1, 2].PutValue(-productB[i]);       // Product B (negative for opposite side)
            }

            // ----- Add a stacked bar chart (horizontal) -----
            // Add chart to the worksheet (rows 0‑15, columns 4‑12)
            int chartIndex = sheet.Charts.Add(ChartType.BarStacked, 0, 4, 15, 12);
            Chart chart = sheet.Charts[chartIndex];

            // Define the data range for the chart (including headers)
            chart.SetChartDataRange("A1:C6", true);

            // Add series manually (required before accessing NSeries items)
            chart.NSeries.Add("B2:B6", true); // Product A series
            chart.NSeries.Add("C2:C6", true); // Product B series

            // Set categories (region names)
            chart.NSeries.CategoryData = "A2:A6";

            // ----- Apply custom colors to each series -----
            // Series 0 – Product A (blue)
            chart.NSeries[0].Area.FillFormat.SolidFill.Color = Color.Blue;
            // Series 1 – Product B (orange)
            chart.NSeries[1].Area.FillFormat.SolidFill.Color = Color.Orange;

            // Ensure the chart layout is calculated before exporting
            chart.Calculate();

            // ----- Export the chart as a PNG image -----
            string imagePath = "tornado_chart.png";
            chart.ToImage(imagePath);   // Extension determines PNG format

            // Optional: save the workbook for reference
            string workbookPath = "tornado_chart.xlsx";
            workbook.Save(workbookPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}