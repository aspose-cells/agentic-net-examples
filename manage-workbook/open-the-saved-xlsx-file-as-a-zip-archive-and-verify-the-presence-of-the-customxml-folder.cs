// Title: Verify the customXml folder in an .xlsx workbook by opening it as a ZipArchive with Aspose.Cells in C#
// Common Searches: how to check for customXml folder inside an .xlsx file using C# and Aspose.Cells | C# open Excel .xlsx as zip to inspect customXml directory | Aspose.Cells verify presence of customXml package part in workbook | read .xlsx package structure with ZipArchive and detect customXml folder | determine if an Excel file contains customXml part using Aspose.Cells C#
// Tags: Aspose.Cells open workbook zip archive | C# inspect Excel package customXml folder | ZipArchive filter entries for customXml | verify customXml part in .xlsx file | read .xlsx internal structure with Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;
using System.IO.Compression;
using System.Linq;

// The example loads an existing .xlsx workbook with Aspose.Cells, then reopens the same file as a ZipArchive to examine its internal entries. It checks whether any entry starts with the "customXml/" path, indicating the presence of the customXml folder, and outputs the verification result while handling potential errors.
class Program
{
    static void Main()
    {
        // Path to the existing .xlsx file
        string filePath = @"C:\Path\To\YourFile.xlsx";

        // Verify that the file exists before attempting to load it
        if (!File.Exists(filePath))
        {
            Console.WriteLine($"File not found: {filePath}");
            return;
        }

        try
        {
            // Load the workbook using Aspose.Cells
            Workbook workbook = new Workbook(filePath);

            // Open the .xlsx file as a zip archive to inspect its internal structure
            using (FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.Read))
            using (ZipArchive zip = new ZipArchive(fs, ZipArchiveMode.Read))
            {
                // Check for the presence of the "customXml" folder
                bool hasCustomXmlFolder = zip.Entries.Any(entry => entry.FullName.StartsWith("customXml/", StringComparison.OrdinalIgnoreCase));

                // Output the verification result
                Console.WriteLine($"customXml folder present: {hasCustomXmlFolder}");
            }
        }
        catch (Exception ex)
        {
            // Handle any errors that occur during loading or processing
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
