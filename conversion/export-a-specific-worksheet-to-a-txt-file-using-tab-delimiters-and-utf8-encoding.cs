// Title: Export a single worksheet to a tab‑delimited UTF‑8 text file with Aspose.Cells for .NET (C#)
// Common Searches: c# aspnet export only one worksheet to tab delimited txt using Aspose.Cells | how to save a specific sheet as UTF-8 text with tab separator in Aspose.Cells | Aspose.Cells TxtSaveOptions example for tab delimited export in C# | export Sheet2 from Excel to UTF-8 txt file using Aspose.Cells .NET | set active worksheet for TxtSaveOptions Aspose.Cells C#
// Tags: Aspose.Cells export worksheet to tab delimited txt | TxtSaveOptions UTF-8 encoding C# | select active sheet for TxtSaveOptions Aspose.Cells | save specific worksheet as text file .NET | configure separator property Aspose.Cells TxtSaveOptions | C# load workbook and export single sheet to txt

using System;
using System.IO;
using System.Text;
using Aspose.Cells;

// The program verifies the source Excel file, loads it with Aspose.Cells, selects the worksheet named "Sheet2", makes it the active sheet, configures TxtSaveOptions to use a tab separator and UTF‑8 encoding, and then saves that worksheet as a tab‑delimited TXT file while handling potential errors.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.txt";

            // Verify that the source workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                return;
            }

            // Load the source workbook
            Workbook workbook = new Workbook(inputPath);

            // Determine the worksheet to export (by name or index)
            Worksheet sheet = workbook.Worksheets["Sheet2"];
            if (sheet == null)
            {
                Console.WriteLine("Error: Worksheet \"Sheet2\" does not exist.");
                return;
            }

            // Activate the selected worksheet so that TxtSaveOptions exports only this sheet
            workbook.Worksheets.ActiveSheetIndex = sheet.Index;

            // Configure TXT save options: tab delimiter and UTF‑8 encoding
            TxtSaveOptions saveOptions = new TxtSaveOptions
            {
                Separator = '\t',               // tab delimiter
                Encoding = Encoding.UTF8        // UTF‑8 encoding
                // ExportColumnHeaders, ExportRowHeaders, and SheetIndex are not required;
                // the active sheet will be exported.
            };

            // Export the selected worksheet to a TXT file
            workbook.Save(outputPath, saveOptions);
            Console.WriteLine($"Worksheet \"{sheet.Name}\" exported successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An unexpected error occurred: {ex.Message}");
        }
    }
}
