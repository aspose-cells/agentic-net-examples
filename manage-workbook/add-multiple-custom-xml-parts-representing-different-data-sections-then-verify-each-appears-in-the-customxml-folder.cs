// Title: Add multiple custom XML parts to an Excel workbook and verify they appear in the customXml folder with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# add multiple custom XML parts to an .xlsx workbook | verify entries in the customXml folder after saving a workbook with Aspose.Cells | list custom XML parts count in a saved Excel file using Aspose.Cells .NET | C# read customXml folder from an .xlsx zip archive | check that custom XML parts are stored in the customXml directory of an Excel package
// Tags: add custom xml parts Aspose.Cells | customXml folder enumeration Aspose.Cells | multiple custom xml sections workbook | verify custom xml part count C# | Aspose.Cells custom xml validation

using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using Aspose.Cells;

// The example creates a new Workbook, adds three distinct custom XML parts as UTF‑8 byte arrays, saves the file as an .xlsx, reloads it to count the parts via the Aspose.Cells API, then opens the saved file as a zip archive to enumerate entries under the 'customXml/' folder and confirms that the folder entry count matches the number of added custom XML parts.
class CustomXmlDemo
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Define custom XML content for different data sections
            string xmlSection1 = @"<Section1><Item>Value1</Item></Section1>";
            string xmlSection2 = @"<Section2><Item>Value2</Item></Section2>";
            string xmlSection3 = @"<Section3><Item>Value3</Item></Section3>";

            // Convert XML strings to byte arrays (UTF-8) – schema data is optional, pass null
            byte[] xmlBytes1 = Encoding.UTF8.GetBytes(xmlSection1);
            byte[] xmlBytes2 = Encoding.UTF8.GetBytes(xmlSection2);
            byte[] xmlBytes3 = Encoding.UTF8.GetBytes(xmlSection3);

            // Add custom XML parts to the workbook
            workbook.CustomXmlParts.Add(xmlBytes1, null);
            workbook.CustomXmlParts.Add(xmlBytes2, null);
            workbook.CustomXmlParts.Add(xmlBytes3, null);

            // Save the workbook to a file
            string filePath = "CustomXmlDemo.xlsx";
            workbook.Save(filePath, SaveFormat.Xlsx);

            // Verify the file exists before loading
            if (!File.Exists(filePath))
                throw new FileNotFoundException($"The file '{filePath}' was not created.");

            // Load the workbook again (optional verification using Aspose.Cells API)
            Workbook loadedWorkbook = new Workbook(filePath);
            int partCount = loadedWorkbook.CustomXmlParts.Count;
            Console.WriteLine($"Number of custom XML parts loaded via Aspose.Cells: {partCount}");

            // Verify that each custom XML part appears in the 'customXml' folder of the package
            using (FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.Read))
            using (ZipArchive zip = new ZipArchive(fs, ZipArchiveMode.Read))
            {
                // Get all entries under the customXml folder
                var customXmlEntries = zip.Entries
                    .Where(e => e.FullName.StartsWith("customXml/", StringComparison.OrdinalIgnoreCase))
                    .ToList();

                Console.WriteLine($"Number of entries in 'customXml' folder: {customXmlEntries.Count}");

                // List the entry names for inspection
                foreach (var entry in customXmlEntries)
                {
                    Console.WriteLine($" - {entry.FullName}");
                }

                // Simple verification: counts should match
                if (customXmlEntries.Count == partCount && partCount == 3)
                {
                    Console.WriteLine("Verification succeeded: All custom XML parts are present in the customXml folder.");
                }
                else
                {
                    Console.WriteLine("Verification failed: Mismatch in custom XML parts count.");
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
