using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

class InMemoryArraySeriesDemo
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // In‑memory double array that will be used as Y‑values for the chart series
            double[] values = new double[] { 12.5, 23.8, 35.1, 44.6, 58.3 };

            // Import the array into the worksheet vertically starting at cell B2 (row index 1, column index 1)
            sheet.Cells.ImportArray(values, 1, 1, true);

            // Determine the address of the imported range (e.g., B2:B6)
            string startCell = "B2";
            string endCell = $"B{1 + values.Length}";
            string dataRange = $"=Sheet1!{startCell}:{endCell}";

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
            Chart chart = sheet.Charts[chartIndex];

            // Add a series to the chart using the imported double array range
            chart.NSeries.Add(dataRange, true);

            // (Optional) Set category labels – here we just use sequential numbers 1..n
            string[] categories = new string[values.Length];
            for (int i = 0; i < categories.Length; i++)
                categories[i] = (i + 1).ToString();

            sheet.Cells.ImportArray(categories, 1, 0, true);
            chart.NSeries.CategoryData = $"=Sheet1!A2:A{1 + values.Length}";

            // Save the workbook
            string outputPath = "InMemoryArraySeries.xlsx";
            workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}