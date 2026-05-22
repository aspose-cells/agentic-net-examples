using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartExample
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // -------------------------------------------------
                // Populate sample data
                // -------------------------------------------------
                // Categories
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["A2"].PutValue("Jan");
                sheet.Cells["A3"].PutValue("Feb");
                sheet.Cells["A4"].PutValue("Mar");

                // Primary series values
                sheet.Cells["B1"].PutValue("Sales");
                sheet.Cells["B2"].PutValue(120);
                sheet.Cells["B3"].PutValue(150);
                sheet.Cells["B4"].PutValue(180);

                // Secondary series values (different scale)
                sheet.Cells["C1"].PutValue("Profit");
                sheet.Cells["C2"].PutValue(3000);
                sheet.Cells["C3"].PutValue(4500);
                sheet.Cells["C4"].PutValue(5000);

                // -------------------------------------------------
                // Add a column chart
                // -------------------------------------------------
                int chartIndex = sheet.Charts.Add(ChartType.Column, 6, 0, 25, 15);
                Chart chart = sheet.Charts[chartIndex];

                // Add primary series (Sales)
                chart.NSeries.Add("B2:B4", true);
                // Add secondary series (Profit)
                chart.NSeries.Add("C2:C4", true);

                // Set category (X‑axis) data
                chart.NSeries.CategoryData = "A2:A4";

                // -------------------------------------------------
                // Configure secondary Y axis
                // -------------------------------------------------
                // Plot the second series on the secondary value axis
                chart.NSeries[1].PlotOnSecondAxis = true;

                // Optional: customize the secondary axis appearance
                Axis secValueAxis = chart.SecondValueAxis;
                secValueAxis.Title.Text = "Profit (USD)";
                secValueAxis.MinValue = 0;
                secValueAxis.MaxValue = 6000;
                secValueAxis.MajorUnit = 1000;

                // -------------------------------------------------
                // Enable cell‑based data labels for the secondary series
                // -------------------------------------------------
                chart.NSeries[1].DataLabels.ShowValue = true;
                chart.NSeries[1].DataLabels.ShowCategoryName = true;

                // -------------------------------------------------
                // Position legend centrally (using Bottom as closest enum value)
                // -------------------------------------------------
                chart.Legend.Position = LegendPositionType.Bottom;
                chart.Legend.IsOverLay = false;

                // -------------------------------------------------
                // Finalize and save the workbook
                // -------------------------------------------------
                chart.Calculate();

                // Save the workbook
                string outputPath = "ChartWithSecondaryAxis_CentralLegend.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}