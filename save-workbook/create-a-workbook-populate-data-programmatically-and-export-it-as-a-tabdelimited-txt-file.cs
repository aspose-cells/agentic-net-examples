// Title: Export a programmatically created workbook to a tab‑delimited TXT file using Aspose.Cells for .NET
// AI Prompts: Write C# code that creates a new Workbook, fills header and data cells, and saves it as a tab‑separated .txt with Aspose.Cells. | Show how to set TxtSaveOptions with a tab character as the separator and UTF‑8 encoding before calling Workbook.Save. | Add logic to verify the output folder exists (creating it if needed) and wrap the save call in a try‑catch that logs any exception.
// Common Searches: asp.net export workbook to tab delimited text file using Aspose.Cells | c# Aspose.Cells TxtSaveOptions tab separator example | how to save Excel data as .txt with tabs in .NET | create workbook, populate rows, and write to tab separated file with Aspose.Cells | ensure output directory exists before saving Aspose.Cells workbook
// Tags: Aspose.Cells TxtSaveOptions tab delimiter | export workbook as tab‑delimited txt | populate worksheet cells programmatically C# | ensure output directory exists .NET | UTF-8 encoding for text export Aspose.Cells

using System;
using System.IO;
using System.Text;
using Aspose.Cells;

// Creates a new Workbook, adds a header row and three data rows, configures TxtSaveOptions with a tab separator and UTF‑8 encoding, ensures the target directory exists, and saves the workbook as a tab‑delimited TXT file.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate header row
            sheet.Cells["A1"].PutValue("ID");
            sheet.Cells["B1"].PutValue("Name");
            sheet.Cells["C1"].PutValue("Score");

            // Populate data rows
            sheet.Cells["A2"].PutValue(1);
            sheet.Cells["B2"].PutValue("Alice");
            sheet.Cells["C2"].PutValue(85);

            sheet.Cells["A3"].PutValue(2);
            sheet.Cells["B3"].PutValue("Bob");
            sheet.Cells["C3"].PutValue(92);

            sheet.Cells["A4"].PutValue(3);
            sheet.Cells["B4"].PutValue("Charlie");
            sheet.Cells["C4"].PutValue(78);

            // Configure TXT save options for tab‑delimited format
            TxtSaveOptions saveOptions = new TxtSaveOptions
            {
                Separator = '\t',                     // Tab delimiter
                Encoding = Encoding.UTF8              // UTF‑8 encoding
            };

            // Define output path
            string outputPath = "ExportedData.txt";

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook as a tab‑delimited TXT file
            workbook.Save(outputPath, saveOptions);
            Console.WriteLine($"Workbook exported successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
