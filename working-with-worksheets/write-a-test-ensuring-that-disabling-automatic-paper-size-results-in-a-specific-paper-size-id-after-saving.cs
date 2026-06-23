using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsDemo
{
    public class PaperSizeTests
    {
        public void DisableAutomaticPaperSize_ShouldRetainSpecificPaperSize()
        {
            try
            {
                // Create a new workbook
                var workbook = new Workbook();

                // Set a specific paper size (A5) which disables automatic paper size
                workbook.Settings.PaperSize = PaperSizeType.PaperA5;

                // Save the workbook to a memory stream
                using (var stream = new MemoryStream())
                {
                    workbook.Save(stream, SaveFormat.Xlsx);
                    stream.Position = 0; // Reset stream position for reading

                    // Load the workbook from the memory stream
                    var loadedWorkbook = new Workbook(stream);
                    var pageSetup = loadedWorkbook.Worksheets[0].PageSetup;

                    // Verify that automatic paper size is disabled
                    if (pageSetup.IsAutomaticPaperSize)
                        throw new Exception("Automatic paper size should be disabled.");

                    // Verify that the paper size matches the expected value (A5)
                    if (pageSetup.PaperSize != PaperSizeType.PaperA5)
                        throw new Exception($"Paper size should be PaperA5 but was {pageSetup.PaperSize}.");
                }

                Console.WriteLine("Test passed: Automatic paper size disabled and PaperA5 retained.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Test failed: {ex.Message}");
            }
        }
    }

    class Program
    {
        static void Main()
        {
            var test = new PaperSizeTests();
            test.DisableAutomaticPaperSize_ShouldRetainSpecificPaperSize();
        }
    }
}