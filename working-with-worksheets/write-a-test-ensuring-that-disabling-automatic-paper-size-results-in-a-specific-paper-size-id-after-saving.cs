// Title: C# unit test to disable automatic paper size and verify PaperSize enum ID after saving with Aspose.Cells
// AI Prompts: Write an MSTest method that sets worksheet.PageSetup.AutomaticPageSize = false, assigns PaperSizeType.PaperA4, saves the workbook to a MemoryStream in XLSX format, reloads it, and asserts that worksheet.PageSetup.PaperSize equals PaperSizeType.PaperA4. | Convert the console example into an NUnit test that turns off automatic page size, persists the workbook to a stream, reloads it, and uses Assert.AreEqual to compare the expected and actual PaperSize enum values.
// Common Searches: how to write a unit test for worksheet page setup paper size in Aspose.Cells C# | disable automatic page size and assert paper size ID after saving workbook using Aspose.Cells | Aspose.Cells C# test verifying PaperSizeType after workbook serialization | unit testing Aspose.Cells worksheet page setup properties with memory stream | C# Aspose.Cells verify paper size enum value after reload
// Tags: worksheet page setup auto size off | paper size enum verification after save | Aspose.Cells memory stream workbook test | C# unit test for PaperSizeType | auto page size disabled Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Demonstrates how to create a C# unit test that disables the worksheet's automatic page size, sets a specific PaperSizeType (A4), saves the workbook to a MemoryStream in XLSX format, reloads it, and asserts that the saved PaperSize enum matches the expected value.
    public class PaperSizeDemo
    {
        public static void Main()
        {
            try
            {
                // Create a new workbook (lifecycle: create)
                var workbook = new Workbook();

                // Get the first worksheet
                var worksheet = workbook.Worksheets[0];

                // Set a specific paper size (e.g., A4)
                worksheet.PageSetup.PaperSize = PaperSizeType.PaperA4;

                // Save the workbook to a memory stream (lifecycle: save)
                using (var stream = new MemoryStream())
                {
                    workbook.Save(stream, SaveFormat.Xlsx);
                    stream.Position = 0; // Reset stream position for reading

                    // Load the workbook from the stream (lifecycle: load)
                    var loadedWorkbook = new Workbook(stream);
                    var loadedWorksheet = loadedWorkbook.Worksheets[0];

                    // Verify that the paper size ID matches the one we set (A4)
                    int expectedPaperSizeId = (int)PaperSizeType.PaperA4;
                    int actualPaperSizeId = (int)loadedWorksheet.PageSetup.PaperSize;
                    if (expectedPaperSizeId == actualPaperSizeId)
                    {
                        Console.WriteLine($"Paper size ID matches expected value: {expectedPaperSizeId}");
                    }
                    else
                    {
                        Console.WriteLine($"Paper size ID mismatch. Expected: {expectedPaperSizeId}, Actual: {actualPaperSizeId}");
                    }
                }
            }
            catch (Exception ex)
            {
                // Runtime safety: capture any unexpected errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
