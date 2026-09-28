// Title: Check for a specific customXml/item1.xml part in a saved XLSX workbook using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an Excel file with Aspose.Cells, saves it to a MemoryStream as XLSX, opens the stream as a ZipArchive, and returns true if the custom XML part is present. | Create a method that iterates over a list of workbook paths, validates the presence of a custom XML part after saving each workbook, and logs the verification result to the console.
// Common Searches: Aspose.Cells C# verify customXml/item1.xml exists after saving workbook | how to read zip entries of a generated .xlsx file using Aspose.Cells .NET | check for customXml/item1.xml in an Excel file created with Aspose.Cells | C# code to confirm customXml/item1.xml inside saved .xlsx archive | validate customXml/item1.xml presence in Aspose.Cells saved workbook
// Tags: Aspose.Cells verify custom XML content in XLSX | C# read XLSX zip entries with ZipArchive | check custom XML after workbook save | validate custom XML presence in saved Excel file | inspect Excel archive for custom XML using .NET

using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using Aspose.Cells;

// The example loads each workbook with Aspose.Cells, saves it to a MemoryStream in XLSX format, opens the stream as a ZipArchive, and checks whether the entry 'customXml/item1.xml' exists, outputting the verification result for every file.
class Program
{
    static void Main()
    {
        // Paths of workbooks to validate
        List<string> workbookPaths = new List<string>
        {
            "Book1.xlsx",
            "Book2.xlsx"
        };

        foreach (string path in workbookPaths)
        {
            // Load the workbook (uses the provided load rule)
            Workbook wb = new Workbook(path);

            // Save to a memory stream (uses the provided save rule)
            using (MemoryStream ms = new MemoryStream())
            {
                wb.Save(ms, SaveFormat.Xlsx);
                ms.Position = 0;

                // Open the saved XLSX as a zip archive
                using (ZipArchive zip = new ZipArchive(ms, ZipArchiveMode.Read, true))
                {
                    // Expected custom XML part entry name
                    const string expectedEntryName = "customXml/item1.xml";

                    bool hasCustomXml = zip.GetEntry(expectedEntryName) != null;

                    Console.WriteLine($"{Path.GetFileName(path)} contains custom XML part '{expectedEntryName}': {hasCustomXml}");
                }
            }
        }
    }
}
