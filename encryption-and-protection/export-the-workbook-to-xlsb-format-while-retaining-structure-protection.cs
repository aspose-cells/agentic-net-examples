// Title: C# – Export Workbook to XLSB with Structure Protection using Aspose.Cells
// Description: Demonstrates how to create or load a workbook, apply password‑protected structure protection, configure XlsbSaveOptions, and save the file as an XLSB document while preserving the protection.
// Keywords: Aspose.Cells XLSB export | workbook structure protection .NET | C# XlsbSaveOptions | protect Excel workbook structure | ExportAllColumnIndexes Aspose | save protected XLSB file
// Common Searches: Aspose.Cells save XLSB with structure protection | C# protect workbook structure and export to XLSB | XlsbSaveOptions ExportAllColumnIndexes example | how to keep workbook protection when saving as XLSB | Aspose.Cells .NET password protected XLSB
// Developer Intent: Save a workbook as an XLSB file while retaining its password‑protected structure.
// Use Cases: Distribute read‑only Excel files that prevent sheet addition, deletion, or renaming. | Create secure template workbooks that maintain layout when shared as XLSB. | Automate generation of compliance‑ready XLSB reports with enforced structure protection.
// AI Prompts: Generate C# code to load an existing workbook, protect its structure with a password, and save it as XLSB using Aspose.Cells. | Show how to disable ExportAllColumnIndexes when exporting to XLSB with Aspose.Cells. | Explain how to programmatically verify that structure protection remains after saving to XLSB.

using System;
using Aspose.Cells;

// Demonstrates how to create or load a workbook, apply password‑protected structure protection, configure XlsbSaveOptions, and save the file as an XLSB document while preserving the protection.
class ExportWorkbookToXlsbWithStructureProtection
{
    static void Main()
    {
        // Create a new workbook (or load an existing one)
        Workbook workbook = new Workbook();

        // Populate some data (optional, just for demonstration)
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Sample Data");
        sheet.Cells["B1"].PutValue(123);

        // Protect the workbook structure with a password
        workbook.Protect(ProtectionType.Structure, "MySecretPassword");

        // Configure XLSB save options
        XlsbSaveOptions saveOptions = new XlsbSaveOptions
        {
            // Ensure column indexes are exported (default is true, set explicitly for clarity)
            ExportAllColumnIndexes = true
        };

        // Save the workbook as XLSB while retaining the structure protection
        workbook.Save("ProtectedWorkbook.xlsb", saveOptions);
    }
}
