// Title: Export a Workbook to XLSB with Structure Protection using Aspose.Cells for .NET (C#)
// Description: Demonstrates how to create or load a workbook, apply structure protection with a password, configure XlsbSaveOptions, and save the file as an XLSB while preserving the protection. The example also shows how to export all column indexes for full fidelity.
// Keywords: Aspose.Cells XLSB export C# | workbook structure protection .NET | XlsbSaveOptions ExportAllColumnIndexes | save protected XLSB Aspose | C# Excel file encryption | global .NET developers | secure Excel export
// Common Searches: Aspose.Cells save XLSB with structure lock | C# protect workbook structure and export to XLSB | ExportAllColumnIndexes option in XlsbSaveOptions | retain workbook protection after XLSB conversion | how to password‑protect Excel workbook using Aspose.Cells
// Developer Intent: Save a workbook as an XLSB file while keeping its structure protection intact.
// Use Cases: Distribute a read‑only XLSB report that prevents adding, deleting, or renaming worksheets. | Provide a template where users can edit cell values but cannot modify the workbook layout. | Archive Excel files in a compact, password‑protected XLSB format for secure long‑term storage.
// AI Prompts: Generate C# code with Aspose.Cells that loads an existing .xlsx, applies structure protection, and saves it as a protected .xlsb. | Explain the purpose of ExportAllColumnIndexes in XlsbSaveOptions and when it should be enabled. | Show how to programmatically confirm that structure protection remains after converting a workbook to XLSB.

using System;
using Aspose.Cells;

namespace AsposeCellsExportXlsb
{
    // Demonstrates how to create or load a workbook, apply structure protection with a password, configure XlsbSaveOptions, and save the file as an XLSB while preserving the protection. The example also shows how to export all column indexes for full fidelity.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook(); // new Workbook("input.xlsx");

            // Add some sample data
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Hello");
            sheet.Cells["B1"].PutValue("World");

            // Protect the workbook structure with a password
            workbook.Protect(ProtectionType.Structure, "MySecretPwd");

            // Configure XLSB save options
            XlsbSaveOptions saveOptions = new XlsbSaveOptions
            {
                // Export all column indexes (optional but often useful)
                ExportAllColumnIndexes = true
            };

            // Save the workbook as XLSB while retaining the structure protection
            workbook.Save("ProtectedWorkbook.xlsb", saveOptions);
        }
    }
}
