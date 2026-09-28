// Title: Add multiple custom XML parts to an Aspose.Cells workbook and list them from the customXml folder in C#
// AI Prompts: Generate C# code that creates a new Workbook, adds two custom XML parts, saves the file as XLSX, and prints the paths of all entries under the customXml directory. | Show how to use Aspose.Cells CustomXmlParts.Add with UTF‑8 byte arrays and then enumerate the customXml folder contents using System.IO.Compression.ZipArchive. | Write a C# snippet that validates the presence of each added custom XML part by checking for its entry name inside the saved Excel package.
// Common Searches: C# Aspose.Cells how to add custom XML sections and confirm they are saved | list entries in customXml directory of a generated XLSX file | use ZipArchive to view customXml parts after saving workbook with Aspose.Cells | check multiple custom XML parts are embedded in Excel package | Aspose.Cells CustomXmlParts.Add example with byte arrays
// Tags: custom XML parts insertion Aspose.Cells C# | customXml folder inspection using ZipArchive | validate custom XML part existence in XLSX | add custom XML without schema data Aspose.Cells | list customXml entries in Excel package

using System;
using System.IO;
using System.IO.Compression;
using System.Text;
using Aspose.Cells;

// Demonstrates adding two custom XML parts to a new workbook, saving it as an XLSX file, and enumerating the entries in the customXml folder of the resulting package using ZipArchive.
class CustomXmlDemo
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Prepare XML data for custom parts
            string xmlSection1 = @"<Section1><Data>Value1</Data></Section1>";
            string xmlSection2 = @"<Section2><Item>Value2</Item></Section2>";

            // Convert XML strings to UTF‑8 byte arrays
            byte[] xmlBytes1 = Encoding.UTF8.GetBytes(xmlSection1);
            byte[] xmlBytes2 = Encoding.UTF8.GetBytes(xmlSection2);

            // Add custom XML parts (schemaData is optional, pass null)
            workbook.CustomXmlParts.Add(xmlBytes1, null);
            workbook.CustomXmlParts.Add(xmlBytes2, null);

            // Save the workbook to a file
            string filePath = "CustomXmlDemo.xlsx";
            workbook.Save(filePath, SaveFormat.Xlsx);

            // Verify that each custom XML part appears in the customXml folder of the package
            if (File.Exists(filePath))
            {
                using (FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.Read))
                using (ZipArchive archive = new ZipArchive(fs, ZipArchiveMode.Read))
                {
                    Console.WriteLine("Custom XML parts found in the package:");
                    foreach (ZipArchiveEntry entry in archive.Entries)
                    {
                        if (entry.FullName.StartsWith("customXml/", StringComparison.OrdinalIgnoreCase))
                        {
                            Console.WriteLine(entry.FullName);
                        }
                    }
                }
            }
            else
            {
                Console.WriteLine($"File not found: {filePath}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
