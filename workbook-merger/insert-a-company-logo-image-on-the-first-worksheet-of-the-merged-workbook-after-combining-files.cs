// Title: Combine multiple Excel files into one workbook and add a company logo to the first sheet using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that loads an array of .xlsx files, merges them into a single Aspose.Cells Workbook, and inserts a PNG logo at cell A1 of the first worksheet. | Demonstrate how to add a picture to a worksheet with Aspose.Cells, set its width and height scale to 50 % and apply a black border. | Write the code to save the merged workbook containing the logo to a new file named MergedWorkbook_WithLogo.xlsx.
// Common Searches: asp.net merge multiple excel workbooks and insert logo using Aspose.Cells | c# Aspose.Cells combine workbooks then add picture to first worksheet | how to place a PNG image at cell A1 after merging Excel files with Aspose.Cells | Aspose.Cells picture scaling and border settings in C# example | save merged workbook with logo using Aspose.Cells C#
// Tags: merge workbooks Aspose.Cells C# | add picture to worksheet Aspose.Cells | insert logo into Excel file Aspose.Cells | scale picture width height Aspose.Cells | apply border to picture Aspose.Cells | save merged workbook with image Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsMergeWithLogo
{
    // Merges several .xlsx files into a single Aspose.Cells workbook, inserts a company logo image at cell A1 of the first worksheet, scales it to 50 % with a black border, and saves the result as MergedWorkbook_WithLogo.xlsx.
    class Program
    {
        static void Main()
        {
            // Paths of the workbooks to be merged
            string[] sourceFiles = new string[]
            {
                "Source1.xlsx",
                "Source2.xlsx",
                // add more files as needed
            };

            // Validate that at least one source file exists
            if (sourceFiles.Length == 0)
            {
                Console.WriteLine("No source files specified.");
                return;
            }

            // Load the first workbook – it will become the destination workbook
            Workbook mergedWorkbook = new Workbook(sourceFiles[0]);

            // Combine the remaining workbooks into the destination workbook
            for (int i = 1; i < sourceFiles.Length; i++)
            {
                Workbook src = new Workbook(sourceFiles[i]);
                mergedWorkbook.Combine(src);
                src.Dispose(); // release resources of the source workbook
            }

            // Insert the company logo on the first worksheet of the merged workbook
            Worksheet firstSheet = mergedWorkbook.Worksheets[0];

            // Path to the logo image file (PNG, JPG, etc.)
            string logoPath = "CompanyLogo.png";

            // Ensure the logo file exists
            if (!File.Exists(logoPath))
            {
                Console.WriteLine($"Logo file not found: {logoPath}");
                mergedWorkbook.Dispose();
                return;
            }

            // Add the picture to the worksheet.
            // Parameters: topRow, leftColumn, fileName
            // Here we place the logo at cell A1 (row 0, column 0)
            int pictureIndex = firstSheet.Pictures.Add(0, 0, logoPath);

            // Optional: adjust picture properties (size, border, etc.)
            Picture logoPicture = firstSheet.Pictures[pictureIndex];
            logoPicture.WidthScale = 50;   // scale width to 50%
            logoPicture.HeightScale = 50;  // scale height to 50%
            logoPicture.BorderLineColor = System.Drawing.Color.Black;
            logoPicture.BorderWeight = 1;

            // Save the merged workbook with the logo
            string outputPath = "MergedWorkbook_WithLogo.xlsx";
            mergedWorkbook.Save(outputPath, SaveFormat.Xlsx);

            // Clean up
            mergedWorkbook.Dispose();

            Console.WriteLine($"Merged workbook saved to: {outputPath}");
        }
    }
}
