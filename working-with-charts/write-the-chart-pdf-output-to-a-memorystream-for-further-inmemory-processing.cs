using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartPdfToMemoryStream
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            worksheet.Cells["A1"].PutValue("Category");
            worksheet.Cells["A2"].PutValue("Apple");
            worksheet.Cells["A3"].PutValue("Banana");
            worksheet.Cells["A4"].PutValue("Cherry");
            worksheet.Cells["B1"].PutValue("Value");
            worksheet.Cells["B2"].PutValue(30);
            worksheet.Cells["B3"].PutValue(45);
            worksheet.Cells["B4"].PutValue(25);

            // Add a column chart to the worksheet
            int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart chart = worksheet.Charts[chartIndex];

            // Set the data source for the chart
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Create a memory stream to hold the PDF output
            using (MemoryStream pdfStream = new MemoryStream())
            {
                // Export the chart to PDF directly into the memory stream
                chart.ToPdf(pdfStream);

                // Reset the stream position if further reading is required
                pdfStream.Position = 0;

                // Example: display the size of the generated PDF
                Console.WriteLine($"Chart PDF generated in memory. Stream length: {pdfStream.Length} bytes");

                // The pdfStream can now be used for further in‑memory processing,
                // such as sending it over a network, attaching to an email, etc.
            }
        }
    }
}