using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsProgressBarDemo
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
                // Prepare data for a progress bar (stacked bar chart)
                // Column A : Category (e.g., tasks)
                // Column B : Full length (background) – set to a constant (e.g., 100)
                // Column C : Actual progress value
                // -------------------------------------------------
                sheet.Cells["A1"].PutValue("Task");
                sheet.Cells["B1"].PutValue("Full");
                sheet.Cells["C1"].PutValue("Progress");

                sheet.Cells["A2"].PutValue("Task 1");
                sheet.Cells["A3"].PutValue("Task 2");
                sheet.Cells["A4"].PutValue("Task 3");
                sheet.Cells["A5"].PutValue("Task 4");

                // Full length – same for all rows (100%)
                for (int row = 2; row <= 5; row++)
                    sheet.Cells[row, 1].PutValue(100);

                // Sample progress values
                sheet.Cells["C2"].PutValue(30);
                sheet.Cells["C3"].PutValue(70);
                sheet.Cells["C4"].PutValue(55);
                sheet.Cells["C5"].PutValue(90);

                // -------------------------------------------------
                // Add a stacked bar chart
                // -------------------------------------------------
                int chartIndex = sheet.Charts.Add(ChartType.BarStacked, 7, 0, 25, 10);
                Chart chart = sheet.Charts[chartIndex];

                // Add the two series (background + progress). 
                // The second argument 'false' indicates that data is in columns.
                chart.NSeries.Add("B2:B5", false); // background series
                chart.NSeries.Add("C2:C5", false); // progress series

                // Set categories (tasks)
                chart.NSeries.CategoryData = "A2:A5";

                // Hide the background series so only the progress part is visible
                Series backgroundSeries = chart.NSeries[0];
                backgroundSeries.IsFiltered = true; // makes the series invisible

                // Customize the visible progress series (e.g., green fill)
                Series progressSeries = chart.NSeries[1];
                progressSeries.Area.ForegroundColor = Color.Green;

                // Optional: Remove gap between bars for a tighter look
                progressSeries.GapWidth = 0;

                // Save the workbook
                string outputPath = "ProgressBarChart.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}