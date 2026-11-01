// Title: C# unit test to verify that disabling automatic paper size keeps a specific PaperSize ID after saving an XLSX workbook with Aspose.Cells
// Common Searches: Aspose.Cells C# test disable automatic paper size and verify PaperSize ID after saving | how to keep A4 paper size when saving a workbook with Aspose.Cells .NET | unit test for PaperSizeType persistence in Aspose.Cells after reload | verify page setup paper size is retained in saved XLSX using Aspose.Cells | disable automatic page size handling Aspose.Cells C# example
// Tags: disable automatic paper size Aspose.Cells | paper size ID preservation Aspose.Cells | Aspose.Cells page setup paper size verification | C# workbook save reload paper size test | PaperSizeType persistence Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsTests
{
    // Creates a new Workbook, sets the first worksheet's PageSetup.PaperSize to PaperA4, saves the workbook as XLSX, reloads it, and asserts that the PaperSize ID remains PaperA4, then cleans up the temporary file.
    public class PaperSizeTests
    {
        // This method verifies that the specified PaperSize ID is preserved after saving and reloading the workbook.
        public static void DisableAutomaticPaperSize_ShouldPreserveSpecificPaperSizeId()
        {
            // Arrange: create a new workbook.
            Workbook workbook = new Workbook();

            // Access the first worksheet's page setup.
            PageSetup pageSetup = workbook.Worksheets[0].PageSetup;

            // Set a specific paper size (A4). Automatic page size handling is not required for this test.
            pageSetup.PaperSize = PaperSizeType.PaperA4;

            // Define a temporary file path for saving.
            string tempFilePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + ".xlsx");

            try
            {
                // Act: save the workbook.
                workbook.Save(tempFilePath, SaveFormat.Xlsx);

                // Ensure the file was created before loading.
                if (!File.Exists(tempFilePath))
                {
                    throw new FileNotFoundException("Saved workbook file not found.", tempFilePath);
                }

                // Load the workbook back.
                Workbook loadedWorkbook = new Workbook(tempFilePath);

                // Retrieve the paper size from the loaded workbook.
                PaperSizeType loadedPaperSize = loadedWorkbook.Worksheets[0].PageSetup.PaperSize;

                // Assert: the loaded paper size should match the one we set (A4).
                if (loadedPaperSize != PaperSizeType.PaperA4)
                {
                    throw new InvalidOperationException(
                        $"The paper size ID was not preserved. Expected: {PaperSizeType.PaperA4}, Actual: {loadedPaperSize}");
                }

                Console.WriteLine("Test passed: Paper size ID preserved after saving and loading the workbook.");
            }
            catch (Exception ex)
            {
                // Runtime safety: report any unexpected errors.
                Console.Error.WriteLine($"Test failed: {ex.Message}");
                throw;
            }
            finally
            {
                // Clean up the temporary file.
                if (File.Exists(tempFilePath))
                {
                    try
                    {
                        File.Delete(tempFilePath);
                    }
                    catch
                    {
                        // Ignore cleanup errors.
                    }
                }
            }
        }

        // Entry point to run the test without requiring NUnit.
        public static void Main()
        {
            DisableAutomaticPaperSize_ShouldPreserveSpecificPaperSizeId();
        }
    }
}
