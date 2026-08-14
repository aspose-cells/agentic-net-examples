// Title: Save a Clean XLSX Workbook and Ensure No printerSettings.bin Using Aspose.Cells for .NET
// Description: Demonstrates how to create or load a Workbook, optionally remove unused styles, save it as an XLSX file with OoxmlSaveOptions, and programmatically verify that the resulting ZIP package does not contain any printerSettings.bin entries.
// Keywords: Aspose.Cells C# | save workbook as XLSX | remove printerSettings.bin | OoxmlSaveOptions ClearData | clean Excel package | verify ZIP archive | .NET Excel export | remove unused styles
// Common Searches: Aspose.Cells remove printerSettings.bin from XLSX | Check XLSX zip for printerSettings.bin C# | Save clean workbook without printer settings Aspose | OoxmlSaveOptions ClearData effect on printer settings | How to verify Excel package contents with Aspose.Cells
// Developer Intent: The developer needs to export a workbook to XLSX, strip unnecessary data such as printer settings, and confirm that the saved file does not embed any printerSettings.bin files.
// Use Cases: Prepare Excel reports for distribution while minimizing file size and removing confidential printer configuration. | Integrate a validation step in CI/CD pipelines that flags XLSX files containing printerSettings.bin. | Automate generation of clean workbooks for third‑party consumption where embedded printer settings are prohibited.
// AI Prompts: Generate C# code with Aspose.Cells that saves a workbook as XLSX and checks the archive for printerSettings.bin entries. | Explain how the ClearData property of OoxmlSaveOptions influences printer settings storage in an XLSX file. | Recommend additional Aspose.Cells save options to reduce file size and ensure no printer configuration data is saved.

using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using Aspose.Cells;

namespace AsposeCellsCleanWorkbook
{
    // Demonstrates how to create or load a Workbook, optionally remove unused styles, save it as an XLSX file with OoxmlSaveOptions, and programmatically verify that the resulting ZIP package does not contain any printerSettings.bin entries.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook();

            // Example data – in real scenario the workbook would already contain data
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Sample");
            sheet.Cells["B2"].PutValue(123);

            // Clean up the workbook: remove unused styles (optional but demonstrates cleaning)
            workbook.RemoveUnusedStyles();

            // Configure OoxmlSaveOptions – ClearData set to false to keep data,
            // other options can be adjusted as needed
            OoxmlSaveOptions saveOptions = new OoxmlSaveOptions
            {
                ClearData = false
            };

            // Save the cleaned workbook as XLSX using the provided Save method
            string outputPath = "cleaned.xlsx";
            workbook.Save(outputPath, saveOptions);

            // Verify that the saved package does not contain any printerSettings.bin files
            bool printerBinExists;
            using (ZipArchive archive = ZipFile.OpenRead(outputPath))
            {
                printerBinExists = archive.Entries
                    .Any(entry => entry.FullName.EndsWith("printerSettings.bin", StringComparison.OrdinalIgnoreCase));
            }

            Console.WriteLine(printerBinExists
                ? "Printer settings .bin file was found inside the package."
                : "No printer settings .bin files exist inside the package.");
        }
    }
}
