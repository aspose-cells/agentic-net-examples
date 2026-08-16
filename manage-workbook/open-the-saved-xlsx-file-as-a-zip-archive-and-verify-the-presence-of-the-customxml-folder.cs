// Title: C# – Check for customXml folder inside an .xlsx workbook using ZipArchive
// Description: Shows how to open an Excel .xlsx file as a ZIP archive with System.IO.Compression, scan its entries for a path that starts with "customXml/", and report whether the customXml directory (which holds custom XML parts) is present.
// Keywords: C# ZipArchive Excel | customXml folder detection | xlsx zip inspection .NET | Aspose.Cells workbook validation | verify custom XML parts | System.IO.Compression | Excel customXml check | US .NET developers | European C# community
// Common Searches: C# check customXml folder in xlsx | how to read .xlsx as zip in .NET | detect custom XML parts in Excel workbook | verify customXml directory using System.IO.Compression | Aspose.Cells verify customXml presence
// Developer Intent: The developer needs to open an .xlsx workbook as a ZIP archive and confirm whether the customXml folder exists.
// Use Cases: Ensure generated workbooks contain required custom XML parts before publishing. | Run automated quality checks on a batch of Excel files to detect missing customXml directories. | Integrate the folder‑presence verification into an Aspose.Cells processing pipeline.
// AI Prompts: Create a reusable C# method that returns true if the customXml folder exists in a given .xlsx file using System.IO.Compression. | Explain how to embed this customXml verification step into an Aspose.Cells workbook creation workflow. | Suggest robust error‑handling and logging strategies when the customXml folder is absent.

using System;
using System.IO;
using System.IO.Compression;

namespace AsposeCellsCustomXmlCheck
{
    // Shows how to open an Excel .xlsx file as a ZIP archive with System.IO.Compression, scan its entries for a path that starts with "customXml/", and report whether the customXml directory (which holds custom XML parts) is present.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the Excel file to be inspected
            string excelPath = "sample.xlsx";

            // Verify that the file exists before proceeding
            if (!File.Exists(excelPath))
            {
                Console.WriteLine($"File not found: {excelPath}");
                return;
            }

            // Open the .xlsx file as a zip archive
            using (FileStream fs = new FileStream(excelPath, FileMode.Open, FileAccess.Read))
            using (ZipArchive archive = new ZipArchive(fs, ZipArchiveMode.Read, leaveOpen: false))
            {
                // Look for an entry that represents the customXml folder
                // In a zip archive, folders are stored as entries ending with a slash
                bool customXmlFolderExists = false;
                foreach (ZipArchiveEntry entry in archive.Entries)
                {
                    // Check if the entry name starts with "customXml/" (case‑insensitive)
                    if (entry.FullName.StartsWith("customXml/", StringComparison.OrdinalIgnoreCase))
                    {
                        customXmlFolderExists = true;
                        break;
                    }
                }

                // Output the verification result
                if (customXmlFolderExists)
                {
                    Console.WriteLine("The customXml folder is present in the workbook.");
                }
                else
                {
                    Console.WriteLine("The customXml folder is NOT present in the workbook.");
                }
            }
        }
    }
}
