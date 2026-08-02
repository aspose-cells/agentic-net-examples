// Title: Detect Font Substitution Warnings When Saving a Workbook to PDF with Aspose.Cells for .NET
// Description: Shows how to apply a non‑existent font to a cell, attach a custom IWarningCallback that records all warnings, export the workbook to PDF, and inspect workbook.Settings.WarningCallback for a FontSubstitution entry.
// Keywords: Aspose.Cells | PDF export | font substitution warning | IWarningCallback | C# | .NET | workbook warnings | missing font detection | PDF conversion testing | warning callback
// Common Searches: Aspose.Cells capture font substitution warning | IWarningCallback example for PDF export .NET | check workbook warnings after saving PDF | detect missing fonts in Aspose.Cells PDF conversion | how to log Aspose.Cells warnings during PDF generation
// Developer Intent: Verify that a FontSubstitution warning is emitted and accessible after converting a workbook to PDF.
// Use Cases: Automated CI test that fails when required fonts are missing during PDF generation. | Centralized logging of missing‑font warnings in large‑scale PDF batch processing. | Programmatic fallback to a default font when a substitution warning is detected.
// AI Prompts: Provide a C# example that registers a custom IWarningCallback to log all warning types during Aspose.Cells PDF export. | Show how to assert the presence of a FontSubstitution warning in an NUnit test for Aspose.Cells PDF saving. | Explain how to retrieve and display warning descriptions from workbook.Settings.WarningCallback after PDF conversion.

using System;
using System.Collections.Generic;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Drawing;

namespace AsposeCellsFontSubstitutionTest
{
    // Custom warning callback that records all warnings raised during rendering/saving
    // Shows how to apply a non‑existent font to a cell, attach a custom IWarningCallback that records all warnings, export the workbook to PDF, and inspect workbook.Settings.WarningCallback for a FontSubstitution entry.
    public class FontSubstitutionWarningCollector : IWarningCallback
    {
        // List to store captured warnings
        public List<WarningInfo> CapturedWarnings { get; } = new List<WarningInfo>();

        // This method is called by Aspose.Cells when a warning occurs
        public void Warning(WarningInfo warningInfo)
        {
            // Store the warning for later inspection
            CapturedWarnings.Add(warningInfo);
        }
    }

    public class Program
    {
        public static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Put some text into a cell and assign a font that is unlikely to exist
            Cell cell = sheet.Cells["A1"];
            cell.PutValue("Text with a non‑existent font");
            Style style = workbook.CreateStyle();
            style.Font.Name = "NonExistentFont";
            cell.SetStyle(style);

            // Instantiate the custom warning collector and assign it to the workbook settings
            FontSubstitutionWarningCollector warningCollector = new FontSubstitutionWarningCollector();
            workbook.Settings.WarningCallback = warningCollector;

            // Configure PDF save options (default options are sufficient for this test)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Save the workbook to PDF; during this process any font substitution warnings will be captured
            workbook.Save("FontSubstitutionTest.pdf", pdfOptions);

            // After saving, examine the collected warnings for a FontSubstitution type
            bool fontSubstitutionWarningFound = false;
            foreach (WarningInfo warning in warningCollector.CapturedWarnings)
            {
                if (warning.WarningType == WarningType.FontSubstitution)
                {
                    fontSubstitutionWarningFound = true;
                    Console.WriteLine($"Font substitution warning captured: {warning.Description}");
                }
            }

            // Output the test result
            if (fontSubstitutionWarningFound)
            {
                Console.WriteLine("Test passed: Font substitution warning was reported.");
            }
            else
            {
                Console.WriteLine("Test failed: No font substitution warning was reported.");
            }
        }
    }
}
