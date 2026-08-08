// Title: Add a Company Logo to the First Sheet of a Merged Workbook with Aspose.Cells (C#)
// Description: C# program that validates three source Excel files, merges them using Aspose.Cells.Combine, inserts a CompanyLogo.png into cells A1‑B5 of the first worksheet, and saves the result as MergedWorkbook_With_Logo.xlsx.
// Keywords: Aspose.Cells merge workbooks C# | insert image Aspose.Cells | add logo to Excel sheet | Pictures.Add example | combine Excel files with branding | C# Excel workbook consolidation
// Common Searches: how to add a logo after merging Excel files Aspose.Cells | Aspose.Cells C# insert picture into first worksheet | merge multiple workbooks and embed image | Aspose.Cells combine workbooks with header image | C# code to place logo on merged Excel sheet
// Developer Intent: Insert a corporate logo into the first worksheet of a workbook created by merging multiple Excel files.
// Use Cases: Produce a single report from departmental spreadsheets with the company logo on the cover page. | Automate client deliverables that combine several worksheets while preserving brand identity. | Generate monthly financial statements by merging files and adding a standardized header image.
// AI Prompts: Show how to position the logo at cell C2 with custom width and height using Aspose.Cells. | Provide code to add the same logo to the header of every worksheet after merging. | Explain how to replace the hard‑coded logo path with a configurable parameter and support PNG, JPEG, and SVG formats.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsMergeWithLogo
{
    // C# program that validates three source Excel files, merges them using Aspose.Cells.Combine, inserts a CompanyLogo.png into cells A1‑B5 of the first worksheet, and saves the result as MergedWorkbook_With_Logo.xlsx.
    class Program
    {
        static void Main()
        {
            try
            {
                // Paths of the workbooks to be merged
                string[] sourceFiles = { "File1.xlsx", "File2.xlsx", "File3.xlsx" };

                // Verify source files exist
                foreach (var file in sourceFiles)
                {
                    if (!File.Exists(file))
                    {
                        Console.WriteLine($"Source file not found: {file}");
                        return;
                    }
                }

                // Path of the company logo image
                string logoPath = "CompanyLogo.png";

                // Load the first workbook which will act as the destination workbook
                Workbook mergedWorkbook = new Workbook(sourceFiles[0]);

                // Combine the remaining workbooks into the destination workbook
                for (int i = 1; i < sourceFiles.Length; i++)
                {
                    Workbook wb = new Workbook(sourceFiles[i]);
                    mergedWorkbook.Combine(wb);
                    wb.Dispose();
                }

                // Insert the logo image on the first worksheet if the file exists
                if (File.Exists(logoPath))
                {
                    Worksheet firstSheet = mergedWorkbook.Worksheets[0];
                    // Add picture occupying cells A1 to B5 (rows 0‑4, columns 0‑1)
                    firstSheet.Pictures.Add(0, 0, 4, 1, logoPath);
                }
                else
                {
                    Console.WriteLine($"Logo file not found: {logoPath}. Skipping logo insertion.");
                }

                // Save the merged workbook with the logo
                string outputPath = "MergedWorkbook_With_Logo.xlsx";
                mergedWorkbook.Save(outputPath, SaveFormat.Xlsx);

                // Clean up
                mergedWorkbook.Dispose();

                Console.WriteLine($"Merged workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
