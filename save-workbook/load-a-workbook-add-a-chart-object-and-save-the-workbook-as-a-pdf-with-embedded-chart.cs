using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;

namespace AsposeCellsChartToPdfDemo
{
    class Program
    {
        static void Main()
        {
            // Load an existing workbook from disk
            string inputFile = "input.xlsx";
            Workbook workbook = new Workbook(inputFile);

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Add sample data for the chart (if the sheet is empty)
            worksheet.Cells["A1"].PutValue("Category");
            worksheet.Cells["B1"].PutValue("Value");
            worksheet.Cells["A2"].PutValue("A");
            worksheet.Cells["B2"].PutValue(10);
            worksheet.Cells["A3"].PutValue("B");
            worksheet.Cells["B3"].PutValue(20);
            worksheet.Cells["A4"].PutValue("C");
            worksheet.Cells["B4"].PutValue(30);

            // Add a column chart to the worksheet
            int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart chart = worksheet.Charts[chartIndex];

            // Set the data source for the chart
            chart.NSeries.Add("B2:B4", true);          // Values
            chart.NSeries.CategoryData = "A2:A4";      // Categories
            chart.Title.Text = "Sample Column Chart";

            // Save the workbook as PDF – the chart will be embedded in the PDF
            string outputPdf = "output.pdf";
            workbook.Save(outputPdf, SaveFormat.Pdf);

            Console.WriteLine($"Workbook saved as PDF with embedded chart to '{outputPdf}'.");
        }
    }
}