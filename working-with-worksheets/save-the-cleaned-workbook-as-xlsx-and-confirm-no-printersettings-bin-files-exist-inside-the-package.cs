using System;
using System.IO;
using System.IO.Compression;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook and add some data
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Sample data");

        // Optional cleanup: remove any unused styles
        workbook.RemoveUnusedStyles();

        // Configure save options for XLSX (OOXML) format
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions
        {
            // Keep the workbook data after saving; we only need to ensure the package is clean
            ClearData = false
        };

        // Save the workbook as an XLSX file
        string outputPath = "CleanedWorkbook.xlsx";
        workbook.Save(outputPath, saveOptions);

        // Verify that the saved XLSX package does not contain any printer‑settings .bin files
        bool printerBinFound = false;
        using (FileStream fs = new FileStream(outputPath, FileMode.Open, FileAccess.Read))
        using (ZipArchive zip = new ZipArchive(fs, ZipArchiveMode.Read))
        {
            foreach (ZipArchiveEntry entry in zip.Entries)
            {
                // Look for entries that are .bin files and contain "printer" in the name
                if (entry.FullName.EndsWith(".bin", StringComparison.OrdinalIgnoreCase) &&
                    entry.FullName.IndexOf("printer", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    printerBinFound = true;
                    Console.WriteLine($"Unexpected printer settings file found: {entry.FullName}");
                }
            }
        }

        Console.WriteLine(printerBinFound
            ? "Printer‑settings .bin files exist inside the package."
            : "No printer‑settings .bin files exist inside the package.");
    }
}