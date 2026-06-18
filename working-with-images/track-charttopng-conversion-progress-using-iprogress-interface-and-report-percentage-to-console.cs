using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace ChartToPngProgressDemo
{
    // Simple IProgress implementation that writes percentage to the console
    class ConsoleProgress : IProgress<int>
    {
        public void Report(int value)
        {
            Console.WriteLine($"Conversion progress: {value}%");
        }
    }

    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("Apple");
            sheet.Cells["A3"].PutValue("Orange");
            sheet.Cells["A4"].PutValue("Banana");

            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["B2"].PutValue(120);
            sheet.Cells["B3"].PutValue(80);
            sheet.Cells["B4"].PutValue(150);

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart chart = sheet.Charts[chartIndex];
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";
            chart.Title.Text = "Fruit Sales";

            // Prepare progress reporter
            IProgress<int> progress = new ConsoleProgress();

            // Report start of conversion
            progress.Report(0);

            // Perform the conversion to PNG
            // The ToImage method saves the chart directly to a file
            chart.ToImage("chart_output.png", ImageType.Png);

            // Report completion of conversion
            progress.Report(100);

            Console.WriteLine("Chart has been successfully saved as PNG.");
        }
    }
}