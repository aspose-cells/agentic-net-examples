using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsTextCrossTypeDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Put a long text into a cell that will exceed the cell width
            sheet.Cells["A1"].PutValue("This is a very long text that will exceed the width of the cell and demonstrate TextCrossType behavior.");
            // Ensure the text is not shrunk to fit
            Style style = sheet.Cells["A1"].GetStyle();
            style.ShrinkToFit = false;
            sheet.Cells["A1"].SetStyle(style);

            // ---------- Image output with different TextCrossType settings ----------
            ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
            {
                ImageType = Aspose.Cells.Drawing.ImageType.Png,
                OnePagePerSheet = true
            };

            // Default behavior (like Excel)
            imgOptions.TextCrossType = TextCrossType.Default;
            SaveSheetAsImage(workbook, imgOptions, "TextCrossType_Default.png");

            // Text crosses cells and keeps existing cell content
            imgOptions.TextCrossType = TextCrossType.CrossKeep;
            SaveSheetAsImage(workbook, imgOptions, "TextCrossType_CrossKeep.png");

            // Text crosses cells and overrides existing cell content
            imgOptions.TextCrossType = TextCrossType.CrossOverride;
            SaveSheetAsImage(workbook, imgOptions, "TextCrossType_CrossOverride.png");

            // Text is confined strictly within the cell boundaries
            imgOptions.TextCrossType = TextCrossType.StrictInCell;
            SaveSheetAsImage(workbook, imgOptions, "TextCrossType_StrictInCell.png");

            // ---------- PDF output with a specific TextCrossType ----------
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Demonstrate CrossKeep for PDF rendering
                TextCrossType = TextCrossType.CrossKeep
            };

            // Save the workbook as PDF using the configured TextCrossType
            workbook.Save("TextCrossType_Pdf_CrossKeep.pdf", pdfOptions);

            Console.WriteLine("Images and PDF have been saved with various TextCrossType settings.");
        }

        // Helper method to render the first sheet as an image using the provided options
        private static void SaveSheetAsImage(Workbook workbook, ImageOrPrintOptions options, string fileName)
        {
            SheetRender renderer = new SheetRender(workbook.Worksheets[0], options);
            renderer.ToImage(0, fileName);
        }
    }
}