using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Saving;
using Aspose.Cells.Utility;

namespace AsposeCellsAddInPdfDemo
{
    public class Program
    {
        public static void Main()
        {
            // 1. Create a new workbook and populate it with sample data
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("Fruits");
            sheet.Cells["A3"].PutValue("Vegetables");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["B2"].PutValue(50);
            sheet.Cells["B3"].PutValue(30);

            // 2. Add a chart that uses the data
            int chartIdx = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart chart = sheet.Charts[chartIdx];
            chart.NSeries.Add("B2:B3", true);
            chart.NSeries.CategoryData = "A2:A3";
            chart.Title.Text = "Sample Chart";

            // 3. Export the chart itself to a separate PDF file
            string chartPdfPath = "ChartOutput.pdf";
            chart.ToPdf(chartPdfPath);
            Console.WriteLine($"Chart exported to PDF: {chartPdfPath}");

            // 4. Convert the entire workbook to PDF
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                EmbedAttachments = true,
                ExportDocumentStructure = true
            };

            string workbookPdfPath = "WorkbookOutput.pdf";
            workbook.Save(workbookPdfPath, pdfOptions);
            Console.WriteLine($"Workbook saved to PDF: {workbookPdfPath}");

            // 5. Demonstrate the utility method that converts an Excel file to PDF
            string tempExcelPath = "TempWorkbook.xlsx";
            workbook.Save(tempExcelPath, SaveFormat.Xlsx);

            string convertedPdfPath = "ConvertedOutput.pdf";
            ConversionUtility.Convert(tempExcelPath, convertedPdfPath);
            Console.WriteLine($"ConversionUtility converted Excel to PDF: {convertedPdfPath}");

            // 6. Clean up temporary files
            if (File.Exists(tempExcelPath)) File.Delete(tempExcelPath);
        }
    }
}