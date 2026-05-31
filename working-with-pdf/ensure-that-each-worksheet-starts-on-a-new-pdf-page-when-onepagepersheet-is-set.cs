using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsExamples
{
    public class OnePagePerSheetDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet and add data
                Worksheet sheet1 = workbook.Worksheets[0];
                sheet1.Name = "FirstSheet";
                for (int i = 0; i < 20; i++)
                {
                    sheet1.Cells[i, 0].PutValue($"Sheet1 Row {i + 1}");
                }

                // Add a second worksheet and add data
                Worksheet sheet2 = workbook.Worksheets.Add("SecondSheet");
                for (int i = 0; i < 30; i++)
                {
                    sheet2.Cells[i, 0].PutValue($"Sheet2 Row {i + 1}");
                }

                // Configure PDF save options to start each worksheet on a new page
                PdfSaveOptions pdfOptions = new PdfSaveOptions
                {
                    OnePagePerSheet = true // Ensures each sheet begins on a new PDF page
                };

                // Save the workbook to PDF using the configured options
                string outputPath = "WorksheetsOnSeparatePages.pdf";
                workbook.Save(outputPath, pdfOptions);
                Console.WriteLine($"PDF saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            OnePagePerSheetDemo.Run();
        }
    }
}