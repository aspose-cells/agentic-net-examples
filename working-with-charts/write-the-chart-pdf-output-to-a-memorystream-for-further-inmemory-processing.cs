using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    public class ChartPdfToMemoryStreamDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Populate sample data for the chart
                worksheet.Cells["A1"].PutValue("Category");
                worksheet.Cells["B1"].PutValue("Value");
                worksheet.Cells["A2"].PutValue("Apple");
                worksheet.Cells["B2"].PutValue(30);
                worksheet.Cells["A3"].PutValue("Banana");
                worksheet.Cells["B3"].PutValue(45);
                worksheet.Cells["A4"].PutValue("Cherry");
                worksheet.Cells["B4"].PutValue(25);

                // Add a column chart to the worksheet
                int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
                Chart chart = worksheet.Charts[chartIndex];

                // Set the data source for the chart
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Export the chart to PDF using a MemoryStream
                using (MemoryStream pdfStream = new MemoryStream())
                {
                    chart.ToPdf(pdfStream);
                    pdfStream.Position = 0; // Reset for reading

                    // Retrieve PDF bytes for further processing
                    byte[] pdfBytes = pdfStream.ToArray();
                    Console.WriteLine($"Chart PDF generated in memory. Stream length: {pdfBytes.Length} bytes");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main()
        {
            ChartPdfToMemoryStreamDemo.Run();
        }
    }
}