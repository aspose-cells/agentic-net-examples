using System;
using System.IO;
using System.IO.Compression;
using System.Linq;

namespace AsposeCellsExamples
{
    class VerifyCustomXmlFolder
    {
        static void Main()
        {
            // Path to the Excel file (xlsx is a zip package)
            string filePath = "output.xlsx";

            // Open the .xlsx file as a zip archive
            using (FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.Read))
            using (ZipArchive zip = new ZipArchive(fs, ZipArchiveMode.Read))
            {
                // Check if any entry belongs to the "customXml" folder
                bool hasCustomXmlFolder = zip.Entries.Any(entry => entry.FullName.StartsWith("customXml/", StringComparison.OrdinalIgnoreCase));

                Console.WriteLine($"CustomXml folder present: {hasCustomXmlFolder}");
            }
        }
    }
}