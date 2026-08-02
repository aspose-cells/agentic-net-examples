// Title: Save a protected workbook as XLSB while preserving structure protection with Aspose.Cells for .NET
// Description: Shows how to create a workbook, add sample data, apply structure protection using a password, set XlsbSaveOptions (ExportAllColumnIndexes), and export the file as a protected XLSB binary workbook.
// Keywords: Aspose.Cells | C# | .NET | XLSB export | structure protection | workbook protection | XlsbSaveOptions | ExportAllColumnIndexes | binary Excel | protected XLSB file
// Common Searches: Aspose.Cells save workbook as XLSB with structure protection | C# protect workbook structure and export to XLSB | XlsbSaveOptions ExportAllColumnIndexes example | keep workbook protection when saving to XLSB | binary Excel file protection Aspose.Cells
// Developer Intent: The developer needs to export a workbook to XLSB format without losing the workbook’s structure protection.
// Use Cases: Distribute a template workbook that must retain its sheet order and be locked against changes, delivered as a binary XLSB file. | Generate financial or HR reports, apply a password‑protected structure lock, and send the protected XLSB to external stakeholders. | Automate a batch job that opens multiple workbooks, adds data, secures the structure, and saves each as a protected XLSB for archival.
// AI Prompts: Write C# code with Aspose.Cells that opens an existing XLSX, applies structure protection using a password, and saves it as a protected XLSB. | Show how to configure XlsbSaveOptions to retain column indexes and keep workbook structure protection when exporting to XLSB. | Explain how to programmatically verify that structure protection remains active after saving a workbook as XLSB with Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsExportXlsb
{
    // Shows how to create a workbook, add sample data, apply structure protection using a password, set XlsbSaveOptions (ExportAllColumnIndexes), and export the file as a protected XLSB binary workbook.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (creation rule)
            Workbook workbook = new Workbook();

            // Populate some sample data
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("ID");
            sheet.Cells["B1"].PutValue("Name");
            sheet.Cells["A2"].PutValue(1);
            sheet.Cells["B2"].PutValue("Alice");
            sheet.Cells["A3"].PutValue(2);
            sheet.Cells["B3"].PutValue("Bob");

            // Protect the workbook structure with a password (protect rule)
            workbook.Protect(ProtectionType.Structure, "mySecretPwd");

            // Create XLSB save options (creation rule)
            XlsbSaveOptions saveOptions = new XlsbSaveOptions
            {
                // Export all column indexes to preserve exact column layout
                ExportAllColumnIndexes = true
            };

            // Save the workbook as XLSB while retaining the protection (save rule)
            workbook.Save("ProtectedWorkbook.xlsb", saveOptions);
        }
    }
}
