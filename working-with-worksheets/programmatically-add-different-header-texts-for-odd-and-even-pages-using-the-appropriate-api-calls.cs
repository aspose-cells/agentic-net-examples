using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsHeaderExample
{
    public class DifferentOddEvenHeaders
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Access the PageSetup object of the worksheet
                PageSetup pageSetup = worksheet.PageSetup;

                // Enable different headers for odd and even pages
                pageSetup.IsHFDiffOddEven = true;

                // Set header for odd pages (default header)
                pageSetup.SetHeader(0, "&LOdd Page Header - Left");   // left
                pageSetup.SetHeader(1, "&CCenter Header - Odd");     // center
                pageSetup.SetHeader(2, "&ROdd Page Header - Right"); // right

                // Set header for even pages
                pageSetup.SetEvenHeader(0, "&LEven Page Header - Left");   // left
                pageSetup.SetEvenHeader(1, "&CEven Page Header - Center"); // center
                pageSetup.SetEvenHeader(2, "&REven Page Header - Right");  // right

                // Define output file path
                string outputPath = "DifferentOddEvenHeaders.xlsx";

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            DifferentOddEvenHeaders.Run();
        }
    }
}