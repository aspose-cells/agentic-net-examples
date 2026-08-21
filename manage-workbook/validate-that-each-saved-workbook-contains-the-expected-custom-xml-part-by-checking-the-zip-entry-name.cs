// Title: C# – Verify a Custom XML Part in a Saved Aspose.Cells Workbook by Scanning the ZIP Package
// Description: Demonstrates how to add a custom XML part to an Aspose.Cells Workbook, save it as an .xlsx file, and confirm the part's presence by opening the file as a ZipArchive and checking for entries under the "customXml/" folder. The example also reloads the workbook to show the API can access the part after saving.
// Keywords: Aspose.Cells custom XML part verification | C# ZipArchive Excel customXml folder | validate custom XML in .xlsx | Aspose.Cells workbook save test | .NET Excel custom XML part check | GitHub Aspose.Cells example | automated Excel package validation | global .NET developers
// Common Searches: How to check if a custom XML part exists in an Excel file saved with Aspose.Cells | C# verify customXml folder in .xlsx using ZipArchive | Aspose.Cells validate custom XML part after saving workbook | Read custom XML entries from Excel package .NET | Unit test for custom XML part in Aspose.Cells workbook
// Developer Intent: Confirm that a workbook saved with Aspose.Cells contains the intended custom XML part by inspecting the underlying ZIP archive entries.
// Use Cases: Automated testing to ensure custom XML parts are embedded before distributing Excel files. | Debugging scenario where developers need to verify the exact package structure of a generated .xlsx. | Continuous‑integration validation that the number of CustomXmlParts matches expectations after a save operation.
// AI Prompts: Generate a C# unit test that creates a workbook with a custom XML part, saves it, and asserts the presence of a "customXml/" entry in the ZIP archive. | Write a method to extract the content of a specific custom XML part from a saved .xlsx using Aspose.Cells and ZipArchive. | Provide code that lists all custom XML part IDs in a loaded workbook and compares them with a predefined list.

using System;
using System.IO;
using System.IO.Compression;
using System.Text;
using Aspose.Cells;
using Aspose.Cells.Markup;

// Demonstrates how to add a custom XML part to an Aspose.Cells Workbook, save it as an .xlsx file, and confirm the part's presence by opening the file as a ZipArchive and checking for entries under the "customXml/" folder. The example also reloads the workbook to show the API can access the part after saving.
class Program
{
    static void Main()
    {
        // Path for the workbook to be saved
        string workbookPath = "CustomXmlDemo.xlsx";

        // Create a new workbook
        Workbook wb = new Workbook();

        // Sample XML data to store in the custom XML part
        string xmlData = "<MyData xmlns=\"http://example.com\"><Value>123</Value></MyData>";
        byte[] dataBytes = Encoding.UTF8.GetBytes(xmlData);

        // Add the custom XML part (no schema in this example)
        int partIndex = wb.CustomXmlParts.Add(dataBytes, null);

        // Optionally set a known ID for later verification
        wb.CustomXmlParts[partIndex].ID = "DemoPartId";

        // Save the workbook using the standard Aspose.Cells Save method
        wb.Save(workbookPath);

        // Validate that the saved workbook contains the expected custom XML part
        bool hasCustomXml = ValidateCustomXmlPart(workbookPath);
        Console.WriteLine($"Custom XML part present in '{workbookPath}': {hasCustomXml}");

        // Load the workbook again to demonstrate that the part can be accessed via the API
        Workbook loadedWb = new Workbook(workbookPath);
        Console.WriteLine($"Custom XML parts count after reload: {loadedWb.CustomXmlParts.Count}");
    }

    // Checks the zip archive of the .xlsx file for an entry under the "customXml/" folder
    static bool ValidateCustomXmlPart(string filePath)
    {
        // .xlsx files are ZIP packages; open it for reading
        using (FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.Read))
        using (ZipArchive archive = new ZipArchive(fs, ZipArchiveMode.Read))
        {
            // Look for any entry whose full name starts with "customXml/"
            foreach (ZipArchiveEntry entry in archive.Entries)
            {
                if (entry.FullName.StartsWith("customXml/", StringComparison.OrdinalIgnoreCase))
                {
                    // Found the custom XML part entry
                    return true;
                }
            }
        }
        // No custom XML part entry found
        return false;
    }
}
