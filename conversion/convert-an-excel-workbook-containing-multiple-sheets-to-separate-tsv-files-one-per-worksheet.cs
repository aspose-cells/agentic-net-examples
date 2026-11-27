// Title: Export each worksheet of an Excel workbook to separate TSV files using Aspose.Cells for .NET
// Common Searches: how to export each sheet of an Excel workbook to a separate TSV file with Aspose.Cells C# | Aspose.Cells C# save individual worksheets as tab‑delimited files | convert multi‑sheet Excel to multiple TSV files programmatically | C# iterate Excel worksheets and generate separate TSV outputs using Aspose | configure TxtSaveOptions for TSV export in Aspose.Cells .NET
// Tags: Aspose.Cells worksheet to TSV export | TxtSaveOptions SaveFormat.Tsv usage | iterate workbook worksheets C# Aspose.Cells | generate separate TSV files per Excel sheet | UTF-8 tab‑delimited file creation with Aspose.Cells

using System;
using System.IO;
using System.Text;
using Aspose.Cells;

// The example loads an Excel workbook, loops through all worksheets, and saves each one as an individual UTF‑8 encoded TSV file by setting TxtSaveOptions with a tab separator and selecting the active sheet before each save.
class ExcelToTsvConverter
{
    static void Main()
    {
        // Path to the source Excel workbook
        string sourcePath = "input.xlsx";

        // Verify that the source file exists
        if (!File.Exists(sourcePath))
        {
            Console.WriteLine($"Source file not found: {sourcePath}");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(sourcePath);

            // Iterate through each worksheet in the workbook
            for (int i = 0; i < workbook.Worksheets.Count; i++)
            {
                try
                {
                    Worksheet sheet = workbook.Worksheets[i];

                    // Build the output TSV file name (using the sheet name)
                    string tsvFileName = $"{sheet.Name}.tsv";

                    // Configure save options for TSV format
                    TxtSaveOptions saveOptions = new TxtSaveOptions(SaveFormat.Tsv)
                    {
                        Separator = '\t',          // Use tab as the delimiter
                        Encoding = Encoding.UTF8   // Ensure UTF-8 encoding
                    };

                    // Set the active sheet to the one we want to export
                    workbook.Worksheets.ActiveSheetIndex = i;

                    // Save the selected worksheet as a TSV file
                    workbook.Save(tsvFileName, saveOptions);
                    Console.WriteLine($"Saved sheet '{sheet.Name}' to {tsvFileName}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error saving sheet index {i}: {ex.Message}");
                }
            }

            Console.WriteLine("Conversion completed.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error during conversion: {ex.Message}");
        }
    }
}
