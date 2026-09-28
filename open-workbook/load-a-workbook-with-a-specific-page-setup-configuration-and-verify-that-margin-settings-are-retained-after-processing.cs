// Title: C# – Load an Excel workbook with predefined page margins using Aspose.Cells for .NET and verify margins remain unchanged after saving
// AI Prompts: Write a C# console program that opens an existing .xlsx file with Aspose.Cells, reads the worksheet PageSetup margin values, modifies a cell, saves the workbook, reloads it, and confirms the margins are identical within a small tolerance. | Generate code that compares TopMargin, BottomMargin, LeftMargin, and RightMargin before and after saving a workbook, then outputs whether the margins were preserved. | Create a reusable C# method that accepts a file path, loads the workbook with Aspose.Cells, performs any processing, and returns true only if the page margins are unchanged after the file is saved.
// Common Searches: Aspose.Cells C# how to ensure page margins are not altered when saving an Excel workbook | verify Excel worksheet margins after modifying cells with Aspose.Cells for .NET | compare original and saved page setup margins using Aspose.Cells in a C# application | check if Aspose.Cells preserves top, bottom, left, and right margins after workbook save | C# code sample to read and validate Excel page margins with Aspose.Cells
// Tags: Aspose.Cells read page setup margins | C# check worksheet margin consistency | Aspose.Cells retain margins on save | Excel page margin validation .NET | Aspose.Cells workbook margin preservation

using System;
using System.IO;
using Aspose.Cells;

// Loads InputWithMargins.xlsx, reads its TopMargin, BottomMargin, LeftMargin, and RightMargin via PageSetup, adds a value to A1, saves as Processed.xlsx, reloads the file, and verifies that all margin values remain unchanged within a 0.001 tolerance.
class Program
{
    static void Main()
    {
        // Path to the source workbook that already has page setup margins configured
        string sourcePath = "InputWithMargins.xlsx";

        // Verify that the source file exists to avoid FileNotFoundException
        if (!File.Exists(sourcePath))
        {
            Console.WriteLine($"Source file not found: {sourcePath}");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(sourcePath);

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Retrieve the original margin settings
            PageSetup pageSetup = sheet.PageSetup;
            double originalTop = pageSetup.TopMargin;
            double originalBottom = pageSetup.BottomMargin;
            double originalLeft = pageSetup.LeftMargin;
            double originalRight = pageSetup.RightMargin;

            // (Optional) Perform any processing on the workbook here
            // For demonstration, we simply add a value to a cell
            sheet.Cells["A1"].PutValue("Test");

            // Save the processed workbook to a new file
            string processedPath = "Processed.xlsx";
            workbook.Save(processedPath);

            // Reload the saved workbook to verify margin retention
            Workbook reloaded = new Workbook(processedPath);
            Worksheet reloadedSheet = reloaded.Worksheets[0];
            PageSetup reloadedSetup = reloadedSheet.PageSetup;

            // Compare margins
            bool marginsRetained =
                Math.Abs(reloadedSetup.TopMargin - originalTop) < 0.001 &&
                Math.Abs(reloadedSetup.BottomMargin - originalBottom) < 0.001 &&
                Math.Abs(reloadedSetup.LeftMargin - originalLeft) < 0.001 &&
                Math.Abs(reloadedSetup.RightMargin - originalRight) < 0.001;

            // Output verification result
            Console.WriteLine(marginsRetained
                ? "Margin settings are retained after processing."
                : "Margin settings have changed.");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors during processing
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
