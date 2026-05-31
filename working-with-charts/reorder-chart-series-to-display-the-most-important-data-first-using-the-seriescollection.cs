using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsSeriesReorder
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

                // Populate sample data for three series
                // Category column
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["A2"].PutValue("Q1");
                sheet.Cells["A3"].PutValue("Q2");
                sheet.Cells["A4"].PutValue("Q3");

                // Series 1 values
                sheet.Cells["B1"].PutValue("Series 1");
                sheet.Cells["B2"].PutValue(10);
                sheet.Cells["B3"].PutValue(20);
                sheet.Cells["B4"].PutValue(30);

                // Series 2 values
                sheet.Cells["C1"].PutValue("Series 2");
                sheet.Cells["C2"].PutValue(40);
                sheet.Cells["C3"].PutValue(15);
                sheet.Cells["C4"].PutValue(25);

                // Series 3 values
                sheet.Cells["D1"].PutValue("Series 3");
                sheet.Cells["D2"].PutValue(5);
                sheet.Cells["D3"].PutValue(35);
                sheet.Cells["D4"].PutValue(45);

                // Add a column chart
                int chartIndex = sheet.Charts.Add(ChartType.Column, 6, 0, 20, 12);
                Chart chart = sheet.Charts[chartIndex];

                // Add all three series at once (by column)
                chart.NSeries.Add("B1:D4", true);

                // Save the workbook
                string outputPath = "SeriesReorder.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                // Log any unexpected errors
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}