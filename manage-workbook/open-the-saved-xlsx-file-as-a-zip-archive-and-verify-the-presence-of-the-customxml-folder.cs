using System;
using System.IO;
using System.IO.Compression;
using Aspose.Cells;

namespace AsposeCellsCustomXmlCheck
{
    class Program
    {
        static void Main()
        {
            // Path to the Excel file that should contain custom XML parts
            string excelPath = "output.xlsx";

            // -----------------------------------------------------------------
            // Load the workbook (using Aspose.Cells lifecycle rule)
            // -----------------------------------------------------------------
            Workbook workbook = new Workbook(excelPath);

            // (Optional) Add a custom XML part if you want to ensure the folder exists
            // This demonstrates the creation of a custom XML part.
            // Comment out if the file already contains custom XML.
            /*
            string xmlData = "<MyData xmlns=\"http://example.com\"/>";
            string xmlSchema = "<?xml version=\"1.0\" encoding=\"UTF-8\"?>"
                             + "<ds:datastoreItem ds:itemID=\"{123}\" xmlns:ds=\"http://schemas.openxmlformats.org/officeDocument/2006/customXml\">"
                             + "<ds:schemaRefs><ds:schemaRef ds:uri=\"http://example.com\"/></ds:schemaRefs>"
                             + "</ds:datastoreItem>";
            workbook.CustomXmlParts.Add(System.Text.Encoding.UTF8.GetBytes(xmlData),
                                        System.Text.Encoding.UTF8.GetBytes(xmlSchema));
            workbook.Save(excelPath);
            */

            // -----------------------------------------------------------------
            // Open the saved .xlsx file as a zip archive
            // -----------------------------------------------------------------
            using (FileStream fs = new FileStream(excelPath, FileMode.Open, FileAccess.Read))
            using (ZipArchive zip = new ZipArchive(fs, ZipArchiveMode.Read, leaveOpen: false))
            {
                // Verify the presence of the "customXml" folder
                bool customXmlFolderExists = false;

                foreach (ZipArchiveEntry entry in zip.Entries)
                {
                    // In a zip archive, a folder is represented by an entry whose name ends with '/'
                    // We check for any entry that starts with "customXml/" (case‑insensitive)
                    if (entry.FullName.StartsWith("customXml/", StringComparison.OrdinalIgnoreCase))
                    {
                        customXmlFolderExists = true;
                        break;
                    }
                }

                Console.WriteLine($"customXml folder present: {customXmlFolderExists}");
            }
        }
    }
}