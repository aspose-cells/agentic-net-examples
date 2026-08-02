// Title: Verify PaperSize ID Persistence After Disabling Automatic Paper Size in Aspose.Cells (.NET)
// Description: A C# example that sets Worksheet.PageSetup.PaperSize to PaperA5 (which turns off automatic sizing), saves the workbook, reloads it, and confirms that IsAutomaticPaperSize remains false and the PaperSize ID is still PaperA5.
// Keywords: Aspose.Cells | .NET | PaperSize | AutomaticPaperSize | PageSetup | Workbook save load | unit test | PaperA5 | paper size persistence
// Common Searches: Aspose.Cells disable automatic paper size | verify paper size after saving workbook | PageSetup.IsAutomaticPaperSize false | PaperSize ID persistence Aspose.Cells | unit test for worksheet page setup
// Developer Intent: Ensure that setting a specific PaperSize disables automatic sizing and that the chosen size ID survives a save‑load cycle.
// Use Cases: Create a workbook, set PageSetup.PaperSize to PaperA5, assert IsAutomaticPaperSize is false, save, reload, and verify PaperSize remains PaperA5. | Repeat the test with other PaperSizeType values (e.g., PaperLetter, PaperLegal) to confirm consistent behavior across formats. | Integrate the verification into CI pipelines to detect regressions in PageSetup serialization.
// AI Prompts: Generate an MSTest method that checks IsAutomaticPaperSize is false and PaperSize ID persists after saving a workbook with Aspose.Cells. | Write a NUnit test that sets Worksheet.PageSetup.PaperSize to PaperA5, saves the file, reloads it, and asserts the PaperSize ID and automatic flag. | Create an xUnit test that validates PaperSize persistence and automatic paper size disabling for Aspose.Cells .NET workbooks.

using System;
using System.Diagnostics;
using System.IO;
using Aspose.Cells;

// A C# example that sets Worksheet.PageSetup.PaperSize to PaperA5 (which turns off automatic sizing), saves the workbook, reloads it, and confirms that IsAutomaticPaperSize remains false and the PaperSize ID is still PaperA5.
class DisableAutomaticPaperSizeTest
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        PageSetup pageSetup = worksheet.PageSetup;

        // Set a specific paper size (A5). This disables automatic paper size.
        pageSetup.PaperSize = PaperSizeType.PaperA5;

        // Verify that automatic paper size is now disabled
        Debug.Assert(!pageSetup.IsAutomaticPaperSize, "IsAutomaticPaperSize should be false after setting PaperSize.");
        Debug.Assert(pageSetup.PaperSize == PaperSizeType.PaperA5, "PaperSize should be PaperA5.");

        // Save the workbook to a temporary file
        string tempFile = Path.Combine(Path.GetTempPath(), "PaperSizeTest.xlsx");
        workbook.Save(tempFile, SaveFormat.Xlsx);

        // Load the workbook back
        Workbook loadedWorkbook = new Workbook(tempFile);
        PageSetup loadedPageSetup = loadedWorkbook.Worksheets[0].PageSetup;

        // Verify that the paper size remains the same and automatic size is still disabled
        if (loadedPageSetup.IsAutomaticPaperSize)
            throw new Exception("Automatic paper size should be disabled after loading.");

        if (loadedPageSetup.PaperSize != PaperSizeType.PaperA5)
            throw new Exception($"Expected PaperSize ID {(int)PaperSizeType.PaperA5}, but got {(int)loadedPageSetup.PaperSize}.");

        Console.WriteLine("Test passed. PaperSize ID after loading: " + (int)loadedPageSetup.PaperSize);
    }
}
