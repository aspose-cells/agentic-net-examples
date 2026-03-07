using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Saving;

namespace AsposeCellsPdfA1aDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and add some sample data
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Product");
            sheet.Cells["B1"].PutValue("Quantity");
            sheet.Cells["A2"].PutValue("Apples");
            sheet.Cells["B2"].PutValue(120);
            sheet.Cells["A3"].PutValue("Oranges");
            sheet.Cells["B3"].PutValue(85);

            // Save the workbook as PDF/A‑1a
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                Compliance = PdfCompliance.PdfA1a
            };

            string pdfPath = "Workbook_PdfA1a.pdf";
            workbook.Save(pdfPath, pdfOptions);
            Console.WriteLine($"Workbook saved as PDF/A‑1a to: {pdfPath}");

            // Render the first page of the workbook to an image (PNG)
            ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
            {
                SaveFormat = SaveFormat.Png
            };

            // Initialize the workbook renderer with the image options
            WorkbookRender renderer = new WorkbookRender(workbook, imgOptions);

            // Render page index 0 (first page) to a PNG file
            string imagePath = "Workbook_Page0.png";
            renderer.ToImage(0, imagePath);
            Console.WriteLine($"First page rendered to image: {imagePath}");
        }
    }
}