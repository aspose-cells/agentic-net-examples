using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace TextCrossTypeScenarios
{
    class Program
    {
        static void Main()
        {
            // Prepare a workbook with a long text that will overflow the cell width
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("This is a very long text that will exceed the width of the cell and demonstrate different TextCrossType behaviors.");
            sheet.Cells["A1"].GetStyle().ShrinkToFit = false; // Ensure text can overflow
            sheet.Cells.SetColumnWidth(0, 10); // Narrow column to force overflow

            // -----------------------------------------------------------------
            // Scenario 1: Export to PNG images with different TextCrossType values
            // -----------------------------------------------------------------
            ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
            {
                ImageType = Aspose.Cells.Drawing.ImageType.Png,
                OnePagePerSheet = true
            };

            foreach (TextCrossType type in Enum.GetValues(typeof(TextCrossType)))
            {
                imgOptions.TextCrossType = type;
                string imgFile = $"Image_TextCrossType_{type}.png";
                SaveSheetAsImage(workbook, imgOptions, imgFile);
                Console.WriteLine($"Saved image with TextCrossType.{type} to {imgFile}");
            }

            // -----------------------------------------------------------------
            // Scenario 2: Export to PDF with different TextCrossType values
            // -----------------------------------------------------------------
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                OnePagePerSheet = true
            };

            foreach (TextCrossType type in Enum.GetValues(typeof(TextCrossType)))
            {
                pdfOptions.TextCrossType = type;
                string pdfFile = $"Pdf_TextCrossType_{type}.pdf";
                workbook.Save(pdfFile, pdfOptions);
                Console.WriteLine($"Saved PDF with TextCrossType.{type} to {pdfFile}");
            }

            // -----------------------------------------------------------------
            // Scenario 3: Using PaginatedSaveOptions directly (same as PdfSaveOptions)
            // -----------------------------------------------------------------
            PaginatedSaveOptions paginatedOptions = new PdfSaveOptions(); // PdfSaveOptions derives from PaginatedSaveOptions
            paginatedOptions.TextCrossType = TextCrossType.CrossOverride;
            paginatedOptions.CheckFontCompatibility = true; // Ensure proper font fallback
            string paginatedPdf = "Paginated_TextCrossType_CrossOverride.pdf";
            workbook.Save(paginatedPdf, paginatedOptions);
            Console.WriteLine($"Saved PDF using PaginatedSaveOptions with CrossOverride to {paginatedPdf}");
        }

        // Helper method to render the first sheet as an image using the supplied options
        private static void SaveSheetAsImage(Workbook workbook, ImageOrPrintOptions options, string fileName)
        {
            SheetRender renderer = new SheetRender(workbook.Worksheets[0], options);
            renderer.ToImage(0, fileName);
        }
    }
}