using System;
using System.Text;
using System.IO;
using System.IO.Compression;
using Aspose.Cells;
using Aspose.Cells.Markup;

namespace AsposeCellsCustomXmlDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (creation rule)
            Workbook workbook = new Workbook();

            // Access the CustomXmlPartCollection
            CustomXmlPartCollection customXmlParts = workbook.CustomXmlParts;

            // Define multiple XML data sections
            string[] xmlDatas = new string[]
            {
                "<Section1><Item>Value1</Item></Section1>",
                "<Section2><Item>Value2</Item></Section2>",
                "<Section3><Item>Value3</Item></Section3>"
            };

            // Optional: define schemas (null in this example)
            byte[] schemaBytes = null;

            // Add each XML part to the workbook
            foreach (string xml in xmlDatas)
            {
                byte[] xmlBytes = Encoding.UTF8.GetBytes(xml);
                // Add method with byte[] data and optional schema (add rule)
                customXmlParts.Add(xmlBytes, schemaBytes);
            }

            // Save the workbook (save rule)
            string outputPath = "MultipleCustomXmlParts.xlsx";
            workbook.Save(outputPath);

            // Verify the number of custom XML parts via the workbook API
            Workbook reloadedWorkbook = new Workbook(outputPath); // load rule
            Console.WriteLine("CustomXmlParts count (API): " + reloadedWorkbook.CustomXmlParts.Count);

            // Verify each part appears in the customXml folder inside the package
            using (ZipArchive archive = ZipFile.OpenRead(outputPath))
            {
                int customXmlEntryCount = 0;
                Console.WriteLine("Entries in customXml folder:");
                foreach (ZipArchiveEntry entry in archive.Entries)
                {
                    if (entry.FullName.StartsWith("customXml/", StringComparison.OrdinalIgnoreCase))
                    {
                        Console.WriteLine("- " + entry.FullName);
                        customXmlEntryCount++;
                    }
                }
                Console.WriteLine("CustomXml folder entry count: " + customXmlEntryCount);
            }
        }
    }
}