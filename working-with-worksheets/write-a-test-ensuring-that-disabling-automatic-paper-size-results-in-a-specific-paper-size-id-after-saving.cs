// Title: C# unit test with Aspose.Cells to verify that a worksheet’s PaperSize setting (A4) is retained after saving and reloading the workbook
// Common Searches: Aspose.Cells verify worksheet paper size after saving workbook C# | C# unit test for persisted PageSetup PaperSize in Excel file using Aspose.Cells | How to check PaperSizeType value after workbook round‑trip with Aspose.Cells | Aspose.Cells disable automatic page size and confirm paper size ID in saved XLSX | MemoryStream save and reload workbook to test PaperSize persistence Aspose.Cells
// Tags: Aspose.Cells verify PaperSize persistence | C# worksheet PageSetup PaperSize test | Aspose.Cells workbook save reload validation | Excel XLSX paper size ID check Aspose.Cells | disable automatic page size Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example creates a workbook, sets the first worksheet's PageSetup PaperSize to A4, saves the file to a MemoryStream in XLSX format, reloads it, and asserts that the PaperSize value remains unchanged, confirming that the setting is persisted across save and load operations.
public class PaperSizeDemo
{
    public static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            var workbook = new Workbook();
            var worksheet = workbook.Worksheets[0];
            var pageSetup = worksheet.PageSetup;

            // Set a specific paper size (A4)
            pageSetup.PaperSize = PaperSizeType.PaperA4;

            // Store the expected paper size for later verification
            var expectedPaperSize = pageSetup.PaperSize;

            // Save the workbook to a memory stream (no file I/O)
            using (var stream = new MemoryStream())
            {
                workbook.Save(stream, SaveFormat.Xlsx);
                stream.Position = 0; // Reset stream position for reading

                // Load the workbook back from the stream
                var loadedWorkbook = new Workbook(stream);
                var loadedWorksheet = loadedWorkbook.Worksheets[0];
                var loadedPageSetup = loadedWorksheet.PageSetup;

                // Verify that the PaperSize persisted correctly
                if (loadedPageSetup.PaperSize != expectedPaperSize)
                {
                    throw new Exception(
                        $"PaperSize mismatch. Expected {expectedPaperSize}, got {loadedPageSetup.PaperSize}.");
                }
            }

            Console.WriteLine("Test passed: PaperSize persisted correctly.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
