using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Charts;

class PdfSample
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate sample data
        worksheet.Cells["A1"].PutValue("Item");
        worksheet.Cells["B1"].PutValue("Quantity");
        worksheet.Cells["A2"].PutValue("Apples");
        worksheet.Cells["B2"].PutValue(10);
        worksheet.Cells["A3"].PutValue("Bananas");
        worksheet.Cells["B3"].PutValue(20);
        worksheet.Cells["A4"].PutValue("Cherries");
        worksheet.Cells["B4"].PutValue(15);

        // Add a column chart based on the data
        int chartIndex = worksheet.Charts.Add(ChartType.Column, 6, 0, 20, 12);
        Chart chart = worksheet.Charts[chartIndex];
        chart.NSeries.Add("B2:B4", true);               // Values
        chart.NSeries.CategoryData = "A2:A4";           // Categories
        chart.Title.Text = "Fruit Quantity";

        // Configure PDF save options
        PdfSaveOptions pdfOptions = new PdfSaveOptions();
        pdfOptions.Compliance = PdfCompliance.PdfA1b;          // Set PDF/A-1b compliance
        pdfOptions.FontEncoding = PdfFontEncoding.Identity;   // Use Identity encoding for fonts
        pdfOptions.Producer = "Aspose.Cells Sample";          // Set custom producer string

        // Save the entire workbook (including the chart) as a PDF file
        workbook.Save("WorkbookWithChart.pdf", pdfOptions);

        // Export only the chart to a separate PDF file
        chart.ToPdf("ChartOnly.pdf");

        Console.WriteLine("PDF files generated successfully.");
    }
}