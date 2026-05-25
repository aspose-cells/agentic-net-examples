using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsExamples
{
    public class Program
    {
        public static void Main()
        {
            OnePagePerSheetAndLimitPagesDemo.Run();
        }
    }

    public class OnePagePerSheetAndLimitPagesDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Populate the workbook with sample data to generate multiple pages
                Worksheet sheet = workbook.Worksheets[0];
                for (int row = 0; row < 200; row++)
                {
                    for (int col = 0; col < 10; col++)
                    {
                        sheet.Cells[row, col].PutValue($"R{row + 1}C{col + 1}");
                    }
                }

                // Configure PDF save options
                PdfSaveOptions pdfOptions = new PdfSaveOptions
                {
                    OnePagePerSheet = true, // Render each worksheet on a single page
                    PageCount = 10          // Limit the output to a maximum of ten pages
                };

                string outputPath = "OnePagePerSheet_LimitedTo10Pages.pdf";

                // Save the workbook as PDF with the specified options
                workbook.Save(outputPath, pdfOptions);
            }
            catch (Exception ex)
            {
                // Log any errors that occur during processing
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}