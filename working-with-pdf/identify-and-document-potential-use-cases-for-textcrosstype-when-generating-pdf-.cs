using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    public class TextCrossTypeUseCasesDemo
    {
        public static void Run()
        {
            // Create a new workbook and obtain the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Insert a long text that will exceed the cell width
            sheet.Cells["A1"].PutValue("This is a very long text that will exceed the width of the cell and demonstrate TextCrossType behavior.");
            // Reduce column width to force overflow
            sheet.Cells.SetColumnWidth(0, 5);

            // ==========================
            // PDF output use cases
            // ==========================
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // 1. Default – behaves like Excel (text may cross into empty neighboring cells)
            pdfOptions.TextCrossType = TextCrossType.Default;
            workbook.Save("TextCross_Default.pdf", pdfOptions);

            // 2. CrossKeep – text crosses cells but does not overwrite existing cell content
            pdfOptions.TextCrossType = TextCrossType.CrossKeep;
            workbook.Save("TextCross_CrossKeep.pdf", pdfOptions);

            // 3. CrossOverride – text crosses cells and overwrites content of crossed cells
            pdfOptions.TextCrossType = TextCrossType.CrossOverride;
            workbook.Save("TextCross_CrossOverride.pdf", pdfOptions);

            // 4. StrictInCell – text is clipped to stay within its own cell boundaries
            pdfOptions.TextCrossType = TextCrossType.StrictInCell;
            workbook.Save("TextCross_StrictInCell.pdf", pdfOptions);

            // ==========================
            // Image output use cases
            // ==========================
            ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
            {
                ImageType = ImageType.Png
            };

            // Helper method to render the first page of the sheet to an image file
            void RenderToImage(string fileName)
            {
                SheetRender sr = new SheetRender(sheet, imgOptions);
                sr.ToImage(0, fileName);
            }

            // 1. Default
            imgOptions.TextCrossType = TextCrossType.Default;
            RenderToImage("TextCross_Default.png");

            // 2. CrossKeep
            imgOptions.TextCrossType = TextCrossType.CrossKeep;
            RenderToImage("TextCross_CrossKeep.png");

            // 3. CrossOverride
            imgOptions.TextCrossType = TextCrossType.CrossOverride;
            RenderToImage("TextCross_CrossOverride.png");

            // 4. StrictInCell
            imgOptions.TextCrossType = TextCrossType.StrictInCell;
            RenderToImage("TextCross_StrictInCell.png");
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            TextCrossTypeUseCasesDemo.Run();
        }
    }
}