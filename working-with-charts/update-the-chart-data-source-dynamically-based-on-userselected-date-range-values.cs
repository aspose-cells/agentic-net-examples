using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

class DynamicChartDataSource
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data: Column A = Date, Column B = Value
            sheet.Cells["A1"].PutValue("Date");
            sheet.Cells["B1"].PutValue("Value");
            DateTime[] dates = {
                new DateTime(2023, 1, 5),
                new DateTime(2023, 2, 10),
                new DateTime(2023, 3, 15),
                new DateTime(2023, 4, 20),
                new DateTime(2023, 5, 25)
            };
            int[] values = { 10, 20, 30, 40, 50 };
            for (int i = 0; i < dates.Length; i++)
            {
                sheet.Cells[i + 2, 0].PutValue(dates[i]);   // Column A
                sheet.Cells[i + 2, 1].PutValue(values[i]); // Column B
            }

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 7, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];

            // User‑selected date range (hard‑coded for demo)
            DateTime userStart = new DateTime(2023, 2, 1);
            DateTime userEnd   = new DateTime(2023, 4, 30);

            // Determine the first and last rows that fall within the selected range
            int firstRow = -1;
            int lastRow = -1;
            for (int row = 2; row <= dates.Length + 1; row++)
            {
                // Retrieve the cell value as DateTime (avoid GetDateTime which may be unavailable)
                object cellObj = sheet.Cells[row, 0].Value;
                if (cellObj is DateTime cellDate)
                {
                    if (cellDate >= userStart && cellDate <= userEnd)
                    {
                        if (firstRow == -1) firstRow = row;
                        lastRow = row;
                    }
                }
            }

            // If matching data exists, set chart data ranges
            if (firstRow != -1)
            {
                string valueRange = $"B{firstRow}:B{lastRow}";
                string categoryRange = $"A{firstRow}:A{lastRow}";

                // Set the chart's data source range (values)
                chart.SetChartDataRange(valueRange, true);

                // Set the category (X‑axis) data range
                chart.NSeries.CategoryData = categoryRange;
            }

            // Display resolved ranges for verification
            Console.WriteLine("Chart Value Range: " + chart.GetChartDataRange());
            Console.WriteLine("Chart Category Range: " + chart.NSeries.CategoryData);

            // Save the workbook
            string outputPath = "DynamicChartDataSource.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}