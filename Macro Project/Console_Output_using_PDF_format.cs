using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate sample data for the chart
        worksheet.Cells["A1"].PutValue("Category");
        worksheet.Cells["A2"].PutValue("Fruits");
        worksheet.Cells["A3"].PutValue("Vegetables");
        worksheet.Cells["B1"].PutValue("Value");
        worksheet.Cells["B2"].PutValue(50);
        worksheet.Cells["B3"].PutValue(30);

        // Add a column chart to the worksheet
        int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
        Chart chart = worksheet.Charts[chartIndex];

        // Set the data source for the chart
        chart.NSeries.Add("B2:B3", true);
        chart.NSeries.CategoryData = "A2:A3";

        // Export the chart to a PDF file
        string chartPdfPath = "ChartOutput.pdf";
        chart.ToPdf(chartPdfPath);
        Console.WriteLine($"Chart exported to PDF: {chartPdfPath}");

        // Save the entire workbook to a PDF file with specific options
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            Compliance = PdfCompliance.PdfA1b,          // Set PDF/A-1b compliance
            Producer = "Aspose.Cells Console Demo"     // Set custom producer string
        };
        string workbookPdfPath = "WorkbookOutput.pdf";
        workbook.Save(workbookPdfPath, pdfOptions);
        Console.WriteLine($"Workbook saved to PDF: {workbookPdfPath}");
    }
}