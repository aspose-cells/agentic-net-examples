using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    public class TextCrossTypeDemo
    {
        public static void Run()
        {
            // Create a new workbook and obtain the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Insert a long text that will exceed the column width
            sheet.Cells["A1"].PutValue("This is a very long text that will cross the cell boundaries when the column is narrow.");

            // Set a narrow column width to force overflow
            sheet.Cells.SetColumnWidth(0, 5);

            // -------------------------------------------------
            // PDF conversion using PdfSaveOptions (inherits TextCrossType from PaginatedSaveOptions)
            // -------------------------------------------------
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Default behavior (like Excel)
            pdfOptions.TextCrossType = TextCrossType.Default;
            workbook.Save("TextCross_Default.pdf", pdfOptions);

            // CrossKeep – text crosses other cells and keeps their content
            pdfOptions.TextCrossType = TextCrossType.CrossKeep;
            workbook.Save("TextCross_CrossKeep.pdf", pdfOptions);

            // CrossOverride – text crosses and overrides other cells
            pdfOptions.TextCrossType = TextCrossType.CrossOverride;
            workbook.Save("TextCross_CrossOverride.pdf", pdfOptions);

            // StrictInCell – text is truncated to the cell width
            pdfOptions.TextCrossType = TextCrossType.StrictInCell;
            workbook.Save("TextCross_StrictInCell.pdf", pdfOptions);

            // -------------------------------------------------
            // Image conversion using ImageOrPrintOptions
            // -------------------------------------------------
            ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
            {
                ImageType = ImageType.Png,
                OnePagePerSheet = true
            };

            // Local function to render and save an image with a specific TextCrossType
            void SaveImage(string fileName, TextCrossType type)
            {
                imgOptions.TextCrossType = type;
                SheetRender render = new SheetRender(sheet, imgOptions);
                render.ToImage(0, fileName);
            }

            // Save images demonstrating each TextCrossType setting
            SaveImage("TextCross_Default.png", TextCrossType.Default);
            SaveImage("TextCross_CrossKeep.png", TextCrossType.CrossKeep);
            SaveImage("TextCross_CrossOverride.png", TextCrossType.CrossOverride);
            SaveImage("TextCross_StrictInCell.png", TextCrossType.StrictInCell);
        }
    }

    public class Program
    {
        public static void Main()
        {
            TextCrossTypeDemo.Run();
        }
    }
}