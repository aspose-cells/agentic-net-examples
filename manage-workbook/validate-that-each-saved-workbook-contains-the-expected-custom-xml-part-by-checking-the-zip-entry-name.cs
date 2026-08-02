using System;
using System.IO;
using System.IO.Compression;
using System.Text;
using Aspose.Cells;
using Aspose.Cells.Markup;

namespace AsposeCellsCustomXmlValidation
{
    class Program
    {
        static void Main()
        {
            // Path for the workbook to be saved
            string workbookPath = "CustomXmlWorkbook.xlsx";

            // ------------------- Create workbook and add custom XML part -------------------
            Workbook wb = new Workbook(); // create workbook (lifecycle rule)

            // Sample XML data and optional schema
            string xmlData = "<MyData xmlns=\"http://example.com\"><Value>123</Value></MyData>";
            byte[] xmlBytes = Encoding.UTF8.GetBytes(xmlData);
            byte[] schemaBytes = null; // no schema needed for this demo

            // Add the custom XML part; Add returns the index of the new part
            int partIndex = wb.CustomXmlParts.Add(xmlBytes, schemaBytes);

            // Optionally set a known ID to retrieve later
            wb.CustomXmlParts[partIndex].ID = "MyCustomXmlPartId";

            // ------------------- Save workbook -------------------
            wb.Save(workbookPath); // save (lifecycle rule)

            // ------------------- Load workbook to verify part count -------------------
            Workbook loadedWb = new Workbook(workbookPath); // load (lifecycle rule)
            Console.WriteLine($"Custom XML parts count after reload: {loadedWb.CustomXmlParts.Count}");

            // ------------------- Validate zip entry name -------------------
            // OOXML stores custom XML parts under the "customXml" folder.
            // We'll open the saved file as a zip archive and look for entries that start with that folder.
            using (FileStream fs = new FileStream(workbookPath, FileMode.Open, FileAccess.Read))
            using (ZipArchive zip = new ZipArchive(fs, ZipArchiveMode.Read))
            {
                bool customXmlEntryFound = false;

                foreach (ZipArchiveEntry entry in zip.Entries)
                {
                    // Check for entries like "customXml/item1.xml"
                    if (entry.FullName.StartsWith("customXml/", StringComparison.OrdinalIgnoreCase) &&
                        entry.FullName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase))
                    {
                        customXmlEntryFound = true;
                        Console.WriteLine($"Found custom XML part entry: {entry.FullName}");
                        break;
                    }
                }

                if (!customXmlEntryFound)
                {
                    Console.WriteLine("Custom XML part entry not found in the workbook package.");
                }
            }
        }
    }
}