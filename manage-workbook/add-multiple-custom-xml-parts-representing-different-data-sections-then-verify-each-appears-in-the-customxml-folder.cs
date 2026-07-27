using System;
using System.Text;
using System.IO;
using System.IO.Compression;
using Aspose.Cells;
using Aspose.Cells.Markup;

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the CustomXmlPartCollection
        CustomXmlPartCollection customXmlParts = workbook.CustomXmlParts;

        // Define XML data for different sections
        string xmlSection1 = "<Section1><Item>Value1</Item></Section1>";
        string xmlSection2 = "<Section2><Item>Value2</Item></Section2>";
        string xmlSection3 = "<Section3><Item>Value3</Item></Section3>";

        // Convert XML strings to byte arrays
        byte[] data1 = Encoding.UTF8.GetBytes(xmlSection1);
        byte[] data2 = Encoding.UTF8.GetBytes(xmlSection2);
        byte[] data3 = Encoding.UTF8.GetBytes(xmlSection3);

        // Add custom XML parts (no schema data)
        int index1 = customXmlParts.Add(data1, null);
        int index2 = customXmlParts.Add(data2, null);
        int index3 = customXmlParts.Add(data3, null);

        // Assign unique IDs (optional, helps identification)
        customXmlParts[index1].ID = Guid.NewGuid().ToString();
        customXmlParts[index2].ID = Guid.NewGuid().ToString();
        customXmlParts[index3].ID = Guid.NewGuid().ToString();

        // Save the workbook to a file
        string outputPath = "MultipleCustomXmlParts.xlsx";
        workbook.Save(outputPath);

        // Verify that each custom XML part appears in the customXml folder of the package
        using (FileStream fs = new FileStream(outputPath, FileMode.Open, FileAccess.Read))
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
}