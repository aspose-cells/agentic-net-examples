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
            // Path for the workbook file
            string workbookPath = "CustomXmlWorkbook.xlsx";

            // ---------- Create ----------
            // Initialize a new workbook
            Workbook wb = new Workbook();

            // Sample XML data and optional schema
            string xmlData = "<MyData xmlns=\"http://example.com\"><Value>123</Value></MyData>";
            byte[] dataBytes = Encoding.UTF8.GetBytes(xmlData);
            byte[] schemaBytes = null; // No schema in this example

            // Add the custom XML part to the workbook
            // The Add method returns the index of the newly added part
            int partIndex = wb.CustomXmlParts.Add(dataBytes, schemaBytes);

            // Optional: set a known ID for later retrieval (not required for zip validation)
            wb.CustomXmlParts[partIndex].ID = Guid.NewGuid().ToString();

            // ---------- Save ----------
            // Save the workbook to disk
            wb.Save(workbookPath);

            // ---------- Load ----------
            // Load the workbook back from the saved file
            Workbook loadedWb = new Workbook(workbookPath);

            // Verify that the custom XML part exists by inspecting the ZIP package
            bool customXmlPartFound = false;
            using (FileStream fs = new FileStream(workbookPath, FileMode.Open, FileAccess.Read))
            using (ZipArchive archive = new ZipArchive(fs, ZipArchiveMode.Read))
            {
                // In OOXML, custom XML parts are stored under the "customXml" folder
                foreach (ZipArchiveEntry entry in archive.Entries)
                {
                    if (entry.FullName.StartsWith("customXml/", StringComparison.OrdinalIgnoreCase) &&
                        entry.FullName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase))
                    {
                        customXmlPartFound = true;
                        Console.WriteLine($"Found custom XML part entry: {entry.FullName}");
                        break;
                    }
                }
            }

            // Output validation result
            if (customXmlPartFound)
            {
                Console.WriteLine("Validation succeeded: Custom XML part is present in the saved workbook.");
            }
            else
            {
                Console.WriteLine("Validation failed: Custom XML part is missing in the saved workbook.");
            }

            // Clean up (optional)
            wb.Dispose();
            loadedWb.Dispose();
        }
    }
}