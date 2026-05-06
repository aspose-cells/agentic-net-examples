using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace AsposeCellsExamples
{
    public class BlankPageHandlingDemo
    {
        public static void Run()
        {
            string sourcePath = "input.xlsx";
            string destPath = "output.pdf";

            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                PrintingPageType = PrintingPageType.IgnoreBlank,
                OutputBlankPageWhenNothingToPrint = false,
                IgnoreError = true
            };

            try
            {
                ConversionUtility.Convert(sourcePath, null, destPath, pdfOptions);
                Console.WriteLine("Conversion completed successfully.");
            }
            catch (CellsException ex) when (ex.Code == ExceptionType.PageSetup)
            {
                Console.WriteLine($"Page setup error during conversion: {ex.Message}");
            }
            catch (CellsException ex)
            {
                Console.WriteLine($"Aspose.Cells error (Code {ex.Code}): {ex.Message}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            BlankPageHandlingDemo.Run();
        }
    }
}