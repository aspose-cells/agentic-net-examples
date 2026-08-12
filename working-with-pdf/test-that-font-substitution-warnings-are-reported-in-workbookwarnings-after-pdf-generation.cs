// Title: Capture Font Substitution Warnings When Exporting a Workbook to PDF with Aspose.Cells for .NET
// Description: Shows how to assign a non‑existent font to a cell, attach a custom IWarningCallback, save the workbook as PDF, and read the FontSubstitution warnings from the callback or workbook.Warnings collection.
// Keywords: Aspose.Cells | PDF export | font substitution warning | IWarningCallback | C# | .NET | missing font detection | warning callback example
// Common Searches: Aspose.Cells capture font substitution warning | C# get PDF export warnings Aspose | how to detect missing fonts in Aspose.Cells PDF | retrieve warnings after workbook.Save as PDF | IWarningCallback usage Aspose.Cells
// Developer Intent: The developer wants to capture and verify font substitution warnings generated during PDF conversion of a workbook.
// Use Cases: Log font substitution events to ensure visual fidelity of generated PDFs. | Fail a CI build when any FontSubstitution warning is reported. | Aggregate warning messages for user notification or audit trails.
// AI Prompts: Write a C# unit test that asserts workbook.Warnings contains a FontSubstitution entry after saving to PDF with Aspose.Cells. | Extend the WarningCollector to store each warning description in a list for later processing. | Explain how to configure Aspose.Cells to treat font substitution warnings as errors instead of informational messages.

using System;
using System.IO;
using Aspose.Cells;

// Shows how to assign a non‑existent font to a cell, attach a custom IWarningCallback, save the workbook as PDF, and read the FontSubstitution warnings from the callback or workbook.Warnings collection.
class FontSubstitutionWarningTest
{
    static void Main()
    {
        try
        {
            // Create a new workbook and access the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add text that uses a font which is unlikely to be installed
            Cell cell = sheet.Cells["A1"];
            cell.PutValue("Text with a missing font");
            Style style = workbook.CreateStyle();
            style.Font.Name = "NonExistentFont";
            cell.SetStyle(style);

            // Set a warning callback to capture warnings during rendering/saving
            var warningCollector = new WarningCollector();
            workbook.Settings.WarningCallback = warningCollector;

            // Save the workbook as PDF (this triggers font substitution processing)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();
            string outputPath = "FontSubstitutionTest.pdf";

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
                Directory.CreateDirectory(outputDir);

            workbook.Save(outputPath, pdfOptions);

            // Output the number of font substitution warnings captured via the callback
            Console.WriteLine($"Font substitution warnings via callback: {warningCollector.FontSubstitutionCount}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }

    // Custom warning callback that records font substitution warnings
    class WarningCollector : IWarningCallback
    {
        public int FontSubstitutionCount { get; private set; }

        public void Warning(WarningInfo warningInfo)
        {
            // Use the updated Type property instead of the obsolete WarningType
            if (warningInfo.Type == ExceptionType.FontSubstitution)
            {
                FontSubstitutionCount++;
                Console.WriteLine($"Warning captured: {warningInfo.Description}");
            }
        }
    }
}
