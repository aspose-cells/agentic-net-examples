// Title: Check for a customXml folder inside a saved .xlsx workbook using Aspose.Cells and C# ZipArchive
// AI Prompts: Generate C# code that loads an existing .xlsx file with Aspose.Cells, saves it, opens the file as a ZipArchive, and returns true when any entry begins with "customXml/". | Create a .NET snippet that enumerates the parts of an Excel package via ZipArchive and determines whether the customXml directory is present.
// Common Searches: C# how to verify that an Excel .xlsx file contains a customXml part | Aspose.Cells detect customXml folder in workbook package programmatically | using System.IO.Compression to list customXml entries in an .xlsx file | check for customXml directory in Office Open XML file with .NET
// Tags: Aspose.Cells inspect xlsx package with ZipArchive | C# verify customXml part in Office Open XML | detect customXml folder in Excel workbook using .NET | zip archive enumeration of Excel file contents | customXml directory presence check in saved workbook

using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using Aspose.Cells;

// Loads an existing workbook, ensures it is saved as .xlsx, opens the file as a ZipArchive, and checks whether any entry resides in the "customXml/" folder, outputting the verification result.
class Program
{
    static void Main()
    {
        // Path to the existing Excel file
        string excelPath = "output.xlsx";

        // Load the workbook (if you need to work with it before verification)
        Workbook workbook = new Workbook(excelPath);

        // Ensure the file is saved (optional if already saved)
        workbook.Save(excelPath, SaveFormat.Xlsx);

        // Open the .xlsx file as a zip archive
        using (FileStream fs = new FileStream(excelPath, FileMode.Open, FileAccess.Read))
        using (ZipArchive archive = new ZipArchive(fs, ZipArchiveMode.Read))
        {
            // Check for any entry that resides in the "customXml" folder
            bool hasCustomXmlFolder = archive.Entries
                .Any(entry => entry.FullName.StartsWith("customXml/", StringComparison.OrdinalIgnoreCase));

            // Output the verification result
            Console.WriteLine(hasCustomXmlFolder
                ? "customXml folder exists in the .xlsx package."
                : "customXml folder NOT found in the .xlsx package.");
        }
    }
}
