using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsExamples
{
    public class PdfCrossStringVerification
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Put a long text in A1 that will exceed the column width
                sheet.Cells["A1"].PutValue(
                    "This is a very long text that should cross the cell boundary and demonstrate the TextCrossType behavior in PDF.");

                // Set a narrow column width to force the text to overflow
                sheet.Cells.SetColumnWidth(0, 5); // Column A width

                // Prepare PDF save options
                PdfSaveOptions pdfOptions = new PdfSaveOptions();

                // Example 1: Default behavior (like Excel)
                pdfOptions.TextCrossType = TextCrossType.Default;
                workbook.Save("CrossString_Default.pdf", pdfOptions);

                // Example 2: Keep the overflow text and keep existing cell contents
                pdfOptions.TextCrossType = TextCrossType.CrossKeep;
                workbook.Save("CrossString_CrossKeep.pdf", pdfOptions);

                // Example 3: Keep the overflow text and override the contents of crossed cells
                pdfOptions.TextCrossType = TextCrossType.CrossOverride;
                workbook.Save("CrossString_CrossOverride.pdf", pdfOptions);

                // Example 4: Strictly keep text inside the cell (truncate overflow)
                pdfOptions.TextCrossType = TextCrossType.StrictInCell;
                workbook.Save("CrossString_StrictInCell.pdf", pdfOptions);

                Console.WriteLine("PDF files generated with different TextCrossType settings.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the console application
    public class Program
    {
        public static void Main(string[] args)
        {
            PdfCrossStringVerification.Run();
        }
    }
}