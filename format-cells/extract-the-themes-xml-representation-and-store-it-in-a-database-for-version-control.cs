// Title: Extract Excel Theme XML (theme1.xml) with Aspose.Cells .NET and Save to Database
// Description: Loads an XLSX workbook using Aspose.Cells, writes it to a memory stream, opens the stream as a ZIP archive, reads the theme XML located at xl/theme/theme1.xml, and passes the XML string to a placeholder method for database insertion or any version‑control system.
// Keywords: Aspose.Cells | C# extract theme XML | theme1.xml | XLSX package | store XML in database | Excel theme version control | read theme from workbook | zip archive | memory stream
// Common Searches: how to read theme1.xml from an xlsx using Aspose.Cells | c# extract excel theme xml and insert into sql server | save workbook theme xml to git repository | asp.net retrieve excel theme for version control | extract theme xml without writing temporary files
// Developer Intent: Retrieve the workbook's theme XML and persist it for version‑control or auditing.
// Use Cases: Archive the extracted theme XML in a Git repo as part of document management. | Insert theme XML into a SQL Server table for change‑tracking and compliance. | Compare theme XML between two workbook revisions to detect style modifications.
// AI Prompts: Generate C# code that uses Aspose.Cells to load an XLSX, locate xl/theme/theme1.xml inside the package, and return the XML as a string. | Create a method that receives the extracted theme XML and stores it in a SQL Server table using a parameterized INSERT command. | Write a robust wrapper that extracts theme XML, handles missing entries, logs errors, and prepares the data for version‑control storage.

using System;
using System.IO;
using System.IO.Compression;
using Aspose.Cells;

namespace AsposeCellsThemeExtraction
{
    // Loads an XLSX workbook using Aspose.Cells, writes it to a memory stream, opens the stream as a ZIP archive, reads the theme XML located at xl/theme/theme1.xml, and passes the XML string to a placeholder method for database insertion or any version‑control system.
    class Program
    {
        // Placeholder for database storage logic
        static void StoreThemeXml(string themeXml)
        {
            // TODO: Implement actual database insertion here.
            // Example: using (var connection = new SqlConnection(connectionString)) { ... }
            Console.WriteLine("Theme XML extracted and ready for storage:");
            Console.WriteLine(themeXml);
        }

        static void Main(string[] args)
        {
            // Path to the source workbook whose theme we want to extract
            string sourcePath = "SourceWorkbook.xlsx";

            // Load the workbook
            Workbook workbook = new Workbook(sourcePath);

            // Save the workbook to a memory stream in XLSX format (the default)
            using (MemoryStream ms = new MemoryStream())
            {
                workbook.Save(ms, SaveFormat.Xlsx);
                ms.Position = 0; // Reset stream position for reading

                // Open the XLSX package as a zip archive
                using (ZipArchive zip = new ZipArchive(ms, ZipArchiveMode.Read, leaveOpen: false))
                {
                    // Theme XML is stored at this path inside the package
                    const string themeEntryPath = "xl/theme/theme1.xml";

                    ZipArchiveEntry themeEntry = zip.GetEntry(themeEntryPath);
                    if (themeEntry != null)
                    {
                        using (StreamReader reader = new StreamReader(themeEntry.Open()))
                        {
                            string themeXml = reader.ReadToEnd();

                            // Store the extracted XML in the database (or any version‑control system)
                            StoreThemeXml(themeXml);
                        }
                    }
                    else
                    {
                        Console.WriteLine("Theme XML not found in the workbook package.");
                    }
                }
            }
        }
    }
}
