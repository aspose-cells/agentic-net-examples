// Title: C# unit test: verify explicit paper size persists after saving with Aspose.Cells
// Description: Creates a workbook, sets Worksheet.PageSetup.PaperSize to PaperA5 (disabling automatic sizing), saves to XLSX, reloads, and asserts IsAutomaticPaperSize is false and PaperSize equals PaperA5. Throws an exception if the values differ.
// Keywords: Aspose.Cells | C# | PaperSize | IsAutomaticPaperSize | unit test | XLSX | PageSetup | PaperA5 | paper size ID 11 | save and reload | regression test
// Common Searches: Aspose.Cells unit test paper size | C# verify PaperSize after save | IsAutomaticPaperSize false Aspose.Cells | persist paper size XLSX Aspose | PageSetup PaperSize test .NET
// Developer Intent: Confirm that disabling automatic paper size stores the chosen PaperSize ID and remains unchanged after workbook serialization.
// Use Cases: Automated regression testing for printing layout consistency | CI validation that custom paper dimensions are retained across save/load cycles | Ensuring PDF generation uses a fixed paper size defined in the workbook | Quality assurance for workbook templates with predefined page‑setup settings
// AI Prompts: Generate an NUnit test that asserts IsAutomaticPaperSize is false and PaperSize equals PaperA5 after saving and loading an XLSX with Aspose.Cells. | Create a MSTest method to verify explicit paper size persistence in a workbook using Aspose.Cells for .NET. | Write a xUnit test checking that PaperSize ID 11 remains after reloading a workbook where automatic paper size was disabled.

using System;
using Aspose.Cells;

// Creates a workbook, sets Worksheet.PageSetup.PaperSize to PaperA5 (disabling automatic sizing), saves to XLSX, reloads, and asserts IsAutomaticPaperSize is false and PaperSize equals PaperA5. Throws an exception if the values differ.
class Program
{
    static void Main()
    {
        // Path for the temporary workbook
        string filePath = "TestPaperSize.xlsx";

        // Create a new workbook and set an explicit paper size (disables automatic size)
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        sheet.PageSetup.PaperSize = PaperSizeType.PaperA5; // PaperSize ID = 11

        // Save the workbook
        workbook.Save(filePath, SaveFormat.Xlsx);

        // Load the saved workbook
        Workbook loadedWorkbook = new Workbook(filePath);
        PageSetup pageSetup = loadedWorkbook.Worksheets[0].PageSetup;

        // Verify that automatic paper size is disabled
        if (pageSetup.IsAutomaticPaperSize)
            throw new Exception("IsAutomaticPaperSize should be false after setting explicit paper size.");

        // Verify that the paper size ID matches the expected value
        if (pageSetup.PaperSize != PaperSizeType.PaperA5)
            throw new Exception($"Expected PaperSize ID {PaperSizeType.PaperA5}, but got {pageSetup.PaperSize}.");

        Console.WriteLine("Test passed: Automatic paper size disabled and PaperSize ID is correct.");
    }
}
