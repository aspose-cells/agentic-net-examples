// Title: Show percentages from worksheet cells in a pie chart’s data labels and resize label shapes using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# how to link pie chart data labels to a cell range for percentages | C# set custom width and height for data label shapes in an Aspose.Cells pie chart | display calculated percentage values in pie chart labels with Aspose.Cells | adjust font and border of data labels in Aspose.Cells chart programmatically | Aspose.Cells example for showing cell‑based percentages in pie chart labels
// Tags: link data labels to cell range Aspose.Cells | pie chart data label shape size Aspose.Cells | display worksheet percentages in chart labels C# | customize data label font border Aspose.Cells | calculate percentages with formulas Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsPieChartDataLabels
{
    // The sample creates a workbook, fills columns with categories, values and percentage formulas, forces calculation, adds a pie chart, links the series data labels to the percentage cells (C2:C4) while hiding default values, sets the label shape to a rectangle with custom width and height ratios, applies font, background, and border styling, and saves the result as PieChartWithCellPercentages.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // ------------------------------------------------------------
                // 1. Populate source data for the pie chart
                // ------------------------------------------------------------
                // Categories
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["A2"].PutValue("Apple");
                sheet.Cells["A3"].PutValue("Banana");
                sheet.Cells["A4"].PutValue("Cherry");

                // Values
                sheet.Cells["B1"].PutValue("Value");
                sheet.Cells["B2"].PutValue(40);
                sheet.Cells["B3"].PutValue(30);
                sheet.Cells["B4"].PutValue(30);

                // ------------------------------------------------------------
                // 2. Calculate percentages in separate cells (C column)
                // ------------------------------------------------------------
                sheet.Cells["C1"].PutValue("Percent");
                sheet.Cells["C2"].Formula = "=B2/SUM($B$2:$B$4)";
                sheet.Cells["C3"].Formula = "=B3/SUM($B$2:$B$4)";
                sheet.Cells["C4"].Formula = "=B4/SUM($B$2:$B$4)";

                // Format as percentage with one decimal place
                Style percentStyle = workbook.CreateStyle();
                percentStyle.Number = 10; // 0.0%
                sheet.Cells["C2:C4"].SetStyle(percentStyle);

                // Force calculation so that linked source values are ready
                workbook.CalculateFormula();

                // ------------------------------------------------------------
                // 3. Add a pie chart
                // ------------------------------------------------------------
                int chartIndex = sheet.Charts.Add(ChartType.Pie, 6, 0, 20, 15);
                Chart pieChart = sheet.Charts[chartIndex];

                // Set data range for the series (values only) and categories
                pieChart.NSeries.Add("B2:B4", true);
                pieChart.NSeries.CategoryData = "A2:A4";

                // ------------------------------------------------------------
                // 4. Configure DataLabels to show percentages from cells
                // ------------------------------------------------------------
                Series series = pieChart.NSeries[0];
                DataLabels dataLabels = series.DataLabels;

                // Show the linked cell range (our calculated percentages)
                dataLabels.ShowCellRange = true;
                dataLabels.LinkedSource = "C2:C4";

                // Hide the default value and percentage that Excel would calculate
                dataLabels.ShowValue = false;
                dataLabels.ShowPercentage = false;

                // ------------------------------------------------------------
                // 5. Adjust shape size and appearance of the data labels
                // ------------------------------------------------------------
                dataLabels.ShapeType = DataLabelShapeType.Rect;

                // Set custom size (fraction of the chart area)
                dataLabels.WidthRatioToChart = 0.12;   // 12% of chart width
                dataLabels.HeightRatioToChart = 0.06;  // 6% of chart height

                // Optional styling
                dataLabels.Font.Size = 10;
                dataLabels.Font.IsBold = true;
                dataLabels.Font.Color = System.Drawing.Color.DarkBlue;
                dataLabels.BackgroundMode = BackgroundMode.Transparent;
                dataLabels.Border.IsVisible = true;
                dataLabels.Border.Color = System.Drawing.Color.Gray;

                // ------------------------------------------------------------
                // 6. Save the workbook
                // ------------------------------------------------------------
                string outputPath = "PieChartWithCellPercentages.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}
