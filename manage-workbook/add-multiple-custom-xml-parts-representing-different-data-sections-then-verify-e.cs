using System;
using System.IO;
using System.IO.Compression;
using System.Text;
using Aspose.Cells;

public class MultipleCustomXmlPartsDemo
{
    public static void Run()
    {
        // 1. Create a new workbook
        Workbook workbook = new Workbook();

        // 2. Define XML data for different sections
        string[] xmlSections = new string[]
        {
            "<Section1><Item>Alpha</Item></Section1>",
            "<Section2><Value>123</Value></Section2>",
            "<Section3><Name>Test</Name></Section3>"
        };

        // 3. Add each XML section as a custom XML part
        foreach (string xml in xmlSections)
        {
            byte[] xmlBytes = Encoding.UTF8.GetBytes(xml);
            // No schema is provided (null)
            workbook.CustomXmlParts.Add(xmlBytes, null);
        }

        // 4. Save the workbook to a file
        string filePath = "MultipleCustomXmlParts.xlsx";
        workbook.Save(filePath);

        // 5. Reload the workbook to ensure persistence
        Workbook loadedWorkbook = new Workbook(filePath);

        // 6. Verify the number of custom XML parts matches the added sections
        int expectedCount = xmlSections.Length;
        int actualCount = loadedWorkbook.CustomXmlParts.Count;
        Console.WriteLine($"Expected custom XML parts: {expectedCount}");
        Console.WriteLine($"Actual custom XML parts:   {actualCount}");

        // 7. Additionally, inspect the package to confirm parts exist in the customXml folder
        int zipCustomXmlCount = 0;
        using (FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.Read))
        using (ZipArchive archive = new ZipArchive(fs, ZipArchiveMode.Read))
        {
            foreach (ZipArchiveEntry entry in archive.Entries)
            {
                // Custom XML parts are stored under the "customXml" directory
                if (entry.FullName.StartsWith("customXml/", StringComparison.OrdinalIgnoreCase))
                {
                    zipCustomXmlCount++;
                }
            }
        }

        Console.WriteLine($"Custom XML parts found in package folder: {zipCustomXmlCount}");

        // 8. Final verification
        if (actualCount == expectedCount && zipCustomXmlCount == expectedCount)
        {
            Console.WriteLine("Verification succeeded: all custom XML parts are present.");
        }
        else
        {
            Console.WriteLine("Verification failed: mismatch in custom XML parts count.");
        }
    }
}

public class Program
{
    public static void Main()
    {
        MultipleCustomXmlPartsDemo.Run();
    }
}