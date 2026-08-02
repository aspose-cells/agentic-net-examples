// Title: Save Aspose.Cells Workbook as XLSX and Verify No printerSettings.bin Files
// Description: Creates a workbook, adds sample data, saves it to XLSX using OoxmlSaveOptions, then opens the file as a ZIP archive to scan for any printerSettings.bin entries and reports the result.
// Keywords: Aspose.Cells save XLSX | printerSettings.bin removal | OoxmlSaveOptions ClearData | inspect XLSX zip contents | .NET Excel package verification
// Common Searches: how to save Aspose.Cells workbook without printerSettings.bin | check XLSX for printerSettings.bin using C# | Aspose.Cells OoxmlSaveOptions prevent printer settings file | verify Excel package is clean of printer settings
// Developer Intent: Save a workbook to XLSX and confirm the generated package does not contain any printerSettings.bin files.
// Use Cases: Distribute clean Excel files that omit printer configuration data. | Automate compliance checks for exported XLSX files before publishing. | Integrate a validation step in CI/CD pipelines to detect unwanted printer settings binaries.
// AI Prompts: Generate C# code with Aspose.Cells that saves a workbook as XLSX and programmatically ensures no printerSettings.bin entries exist. | Explain the effect of OoxmlSaveOptions.ClearData on printerSettings.bin inclusion in the saved XLSX. | Suggest alternative methods to strip printerSettings.bin from an existing XLSX using Aspose.Cells APIs.

using System;
using System.IO;
using System.IO.Compression;
using Aspose.Cells;

namespace AsposeCellsCleanSaveDemo
{
    // Creates a workbook, adds sample data, saves it to XLSX using OoxmlSaveOptions, then opens the file as a ZIP archive to scan for any printerSettings.bin entries and reports the result.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and add some sample data
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Sample");
            sheet.Cells["B2"].PutValue(123);

            // Configure OoxmlSaveOptions for XLSX format
            OoxmlSaveOptions saveOptions = new OoxmlSaveOptions
            {
                // Ensure the workbook is not cleared after saving
                ClearData = false
            };

            // Define the output file path
            string outputPath = "CleanedWorkbook.xlsx";

            // Save the workbook using the provided Save method with options
            workbook.Save(outputPath, saveOptions);

            // Verify that the saved XLSX package does not contain any printerSettings.bin files
            bool printerSettingsFound = false;
            using (FileStream fs = new FileStream(outputPath, FileMode.Open, FileAccess.Read))
            using (ZipArchive zip = new ZipArchive(fs, ZipArchiveMode.Read))
            {
                foreach (ZipArchiveEntry entry in zip.Entries)
                {
                    if (entry.FullName.EndsWith("printerSettings.bin", StringComparison.OrdinalIgnoreCase))
                    {
                        printerSettingsFound = true;
                        break;
                    }
                }
            }

            // Output verification result
            if (printerSettingsFound)
                Console.WriteLine("Warning: printerSettings.bin file was found inside the XLSX package.");
            else
                Console.WriteLine("Success: No printerSettings.bin files exist inside the XLSX package.");
        }
    }
}
