using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    public class ExcelToPdfIgnoreErrorDemo
    {
        public static void Run()
        {
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            worksheet.Cells["A1"].PutValue("Category");
            worksheet.Cells["A2"].PutValue("A");
            worksheet.Cells["A3"].PutValue("B");
            worksheet.Cells["B1"].PutValue("Value");
            worksheet.Cells["B2"].PutValue(10);
            worksheet.Cells["B3"].PutValue(20);

            int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 15, 8);
            Chart chart = worksheet.Charts[chartIndex];
            chart.NSeries.Add("B2:B3", true);
            chart.NSeries.CategoryData = "A2:A3";

            PdfSaveOptions pdfOptions = new PdfSaveOptions();
            pdfOptions.IgnoreError = true;

            workbook.Save("ExcelToPdf_IgnoredErrors.pdf", pdfOptions);

            Console.WriteLine("PDF saved with rendering errors ignored.");
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            ExcelToPdfIgnoreErrorDemo.Run();
        }
    }
}