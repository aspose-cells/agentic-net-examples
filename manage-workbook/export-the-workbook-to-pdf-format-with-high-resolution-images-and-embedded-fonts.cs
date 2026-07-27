using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;

class ExportWorkbookToPdf
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate some sample data
        worksheet.Cells["A1"].PutValue("Category");
        worksheet.Cells["A2"].PutValue("Fruits");
        worksheet.Cells["A3"].PutValue("Vegetables");
        worksheet.Cells["B1"].PutValue("Value");
        worksheet.Cells["B2"].PutValue(50);
        worksheet.Cells["B3"].PutValue(30);

        // Add a chart so the PDF will contain an image that can be rendered at high resolution
        int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
        Chart chart = worksheet.Charts[chartIndex];
        chart.NSeries.Add("B2:B3", true);
        chart.NSeries.CategoryData = "A2:A3";
        chart.Title.Text = "Sample Chart";

        // Configure PDF save options
        PdfSaveOptions pdfOptions = new PdfSaveOptions();

        // Embed standard Windows fonts into the PDF
        pdfOptions.EmbedStandardWindowsFonts = true;

        // Set a default font to be used when a cell's font is missing or unsupported
        pdfOptions.DefaultFont = "Arial";

        // Ensure the workbook's default font is checked for Unicode characters
        pdfOptions.CheckWorkbookDefaultFont = true;

        // Increase image resolution to 300 DPI and set JPEG quality to 100 (maximum)
        pdfOptions.SetImageResample(300, 100);

        // Save the workbook as a PDF with the specified options
        workbook.Save("output_highres.pdf", pdfOptions);
    }
}