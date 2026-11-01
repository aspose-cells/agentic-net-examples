// Title: Insert a PNG company logo into the first worksheet of a merged Excel workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that loads several .xlsx files with Aspose.Cells, combines them into a single workbook, and embeds a PNG logo at cell A1 of the first worksheet, setting its width and height. | Provide a C# snippet that verifies the presence of a logo file before calling Worksheet.Pictures.Add in an Aspose.Cells merged workbook. | Show how to adjust the dimensions of a Picture object after inserting it into a worksheet created by merging Excel workbooks with Aspose.Cells.
// Common Searches: C# Aspose.Cells add company logo to first sheet after merging Excel files | insert PNG image into merged workbook's first worksheet using Aspose.Cells .NET | how to resize a picture added to an Aspose.Cells worksheet after combining workbooks
// Tags: Aspose.Cells merge workbooks logo insertion | C# insert PNG picture into Excel worksheet | Aspose.Cells picture resizing on worksheet | verify image file existence Aspose.Cells | add picture to first sheet of merged workbook

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example loads three source .xlsx files, uses the first as the base workbook, copies worksheets from the others, checks for a PNG logo file, inserts the logo at cell A1 of the first worksheet with a size of 150 × 75 pixels, and saves the combined workbook as MergedWorkbook_With_Logo.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Paths of the workbooks to be merged
            List<string> sourceFiles = new List<string>
            {
                "File1.xlsx",
                "File2.xlsx",
                "File3.xlsx"
            };

            // Verify that source files exist
            foreach (var file in sourceFiles)
            {
                if (!File.Exists(file))
                {
                    Console.WriteLine($"Source file not found: {file}");
                    return;
                }
            }

            // Load the first workbook – it will serve as the base for the merged result
            Workbook mergedWorkbook = new Workbook(sourceFiles[0]);

            // Merge the remaining workbooks into the base workbook
            for (int i = 1; i < sourceFiles.Count; i++)
            {
                // Load the next workbook
                Workbook wb = new Workbook(sourceFiles[i]);

                // Copy each worksheet from the source workbook into the merged workbook
                foreach (Worksheet ws in wb.Worksheets)
                {
                    // Add a copy of the worksheet to the merged workbook using its name
                    mergedWorkbook.Worksheets.AddCopy(ws.Name);
                }
            }

            // Insert the company logo on the first worksheet of the merged workbook
            Worksheet firstSheet = mergedWorkbook.Worksheets[0];

            // Verify that the logo image exists
            string logoPath = "CompanyLogo.png";
            if (!File.Exists(logoPath))
            {
                Console.WriteLine($"Logo file not found: {logoPath}");
                // Continue without adding the logo
            }
            else
            {
                // Add the picture; the picture will be placed at cell A1 (row 0, column 0)
                int pictureIndex = firstSheet.Pictures.Add(0, 0, logoPath);

                // Adjust picture size if needed
                Picture logo = firstSheet.Pictures[pictureIndex];
                logo.Width = 150;   // width in pixels
                logo.Height = 75;   // height in pixels
            }

            // Save the merged workbook with the logo
            string outputPath = "MergedWorkbook_With_Logo.xlsx";
            mergedWorkbook.Save(outputPath);
            Console.WriteLine($"Merged workbook saved to: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
