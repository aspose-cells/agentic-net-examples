// Title: Insert Company Logo into First Sheet After Merging Workbooks with Aspose.Cells (.NET)
// Description: C# code that merges multiple Excel files using Workbook.Combine, guarantees a worksheet exists, and adds a PNG logo to the A1:E5 range of the first sheet with Worksheet.Pictures.Add. Includes checks for missing files, logo presence, and error handling, then saves the result as MergedWithLogo.xlsx.
// Keywords: Aspose.Cells merge workbooks C# | add image to Excel sheet Aspose | insert logo first worksheet | Workbook.Combine example | Worksheet.Pictures.Add C# | handle missing files Aspose.Cells | save merged workbook with logo
// Common Searches: how to merge several Excel files and add a logo using Aspose.Cells | Aspose.Cells insert picture into first worksheet after combine | C# combine workbooks and place image at specific cells | add company logo to merged Excel workbook Aspose | error handling when merging Excel files with Aspose.Cells
// Developer Intent: Merge multiple Excel workbooks into one file and embed a company logo on the first worksheet.
// Use Cases: Create a consolidated monthly report and brand the cover sheet with the corporate logo. | Generate a master presentation workbook that automatically includes the logo after merging departmental files. | Automate internal data aggregation while ensuring every combined workbook starts with consistent branding.
// AI Prompts: Show how to move the logo to the top‑right corner of the first sheet instead of A1:E5. | Provide code to add a semi‑transparent watermark image to every worksheet after merging. | Explain how to support SVG, JPEG, and BMP formats and scale the logo proportionally when using Worksheet.Pictures.Add.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// C# code that merges multiple Excel files using Workbook.Combine, guarantees a worksheet exists, and adds a PNG logo to the A1:E5 range of the first sheet with Worksheet.Pictures.Add. Includes checks for missing files, logo presence, and error handling, then saves the result as MergedWithLogo.xlsx.
class InsertLogoAfterMerge
{
    static void Main()
    {
        // Files that need to be merged
        string[] filesToMerge = { "File1.xlsx", "File2.xlsx", "File3.xlsx" };

        // Create an empty workbook that will hold the merged result
        Workbook mergedWorkbook = new Workbook();

        // Load each source workbook and combine it into the destination workbook
        foreach (string filePath in filesToMerge)
        {
            try
            {
                if (!File.Exists(filePath))
                {
                    Console.WriteLine($"Warning: File not found '{filePath}'. Skipping.");
                    continue;
                }

                // Load source workbook
                Workbook sourceWorkbook = new Workbook(filePath);

                // Combine source workbook into the merged workbook
                mergedWorkbook.Combine(sourceWorkbook);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing '{filePath}': {ex.Message}");
            }
        }

        // Ensure there is at least one worksheet
        if (mergedWorkbook.Worksheets.Count == 0)
        {
            mergedWorkbook.Worksheets.Add();
        }

        // Get the first worksheet of the merged workbook
        Worksheet firstWorksheet = mergedWorkbook.Worksheets[0];

        // Path to the company logo image
        string logoPath = "CompanyLogo.png";

        // Insert the logo picture if the file exists
        if (File.Exists(logoPath))
        {
            // The picture will occupy rows 0‑4 and columns 0‑4 (A1:E5 area)
            firstWorksheet.Pictures.Add(0, 0, 4, 4, logoPath);
        }
        else
        {
            Console.WriteLine($"Warning: Logo file '{logoPath}' not found. Skipping logo insertion.");
        }

        // Save the merged workbook with the inserted logo
        try
        {
            mergedWorkbook.Save("MergedWithLogo.xlsx", SaveFormat.Xlsx);
            Console.WriteLine("Merged workbook saved as 'MergedWithLogo.xlsx'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error saving merged workbook: {ex.Message}");
        }
    }
}
