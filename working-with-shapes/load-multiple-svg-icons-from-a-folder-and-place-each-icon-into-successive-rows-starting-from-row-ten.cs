// Title: Batch load SVG icons into Excel rows starting at row 10 with Aspose.Cells (C#)
// Description: C# program that scans a folder for *.svg files, creates a new workbook, and inserts each SVG as a picture into column A of the first worksheet. Icons are placed in consecutive rows beginning with Excel row 10, and the workbook is saved as an .xlsx file.
// Keywords: Aspose.Cells C# add SVG picture | load SVG icons batch | insert SVG into Excel rows | Aspose.Cells shape collection | add picture from stream | Excel row 10 | batch import SVG files | C# file system iterate SVG | save workbook as xlsx
// Common Searches: Aspose.Cells add SVG picture to worksheet | C# load multiple SVG files into Excel | Insert SVG icons into successive rows | Batch import SVG icons with Aspose.Cells | Place SVG images in column A starting at row 10
// Developer Intent: Read every SVG file in a directory and place it as a picture in a new row of the worksheet, beginning with row 10.
// Use Cases: Generate an icon catalog for a design system directly in Excel. | Create an inventory sheet where each product is represented by its SVG icon. | Prepare a visual asset list for documentation or marketing materials. | Automate a slide‑deck data source by mapping SVG thumbnails to rows. | Build a printable price list that includes product icons.
// AI Prompts: Provide C# Aspose.Cells code that reads all *.svg files from a directory and adds each as a picture to column A, beginning at row 10, with automatic row increment. | Extend the sample to resize each SVG to 64 px height while preserving aspect ratio. | Add logging and exception handling for missing or corrupted SVG files, and skip files that cannot be loaded. | Modify the code to place icons in a user‑specified column instead of column A. | Create a PowerShell wrapper that invokes the compiled program for a given folder path.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// C# program that scans a folder for *.svg files, creates a new workbook, and inserts each SVG as a picture into column A of the first worksheet. Icons are placed in consecutive rows beginning with Excel row 10, and the workbook is saved as an .xlsx file.
class LoadSvgIconsIntoWorksheet
{
    static void Main()
    {
        try
        {
            // Path to the folder containing SVG files
            string svgFolderPath = @"C:\Icons";

            // Verify that the folder exists to avoid DirectoryNotFoundException
            if (!Directory.Exists(svgFolderPath))
            {
                Console.WriteLine($"Folder not found: {svgFolderPath}");
                return;
            }

            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            ShapeCollection shapes = worksheet.Shapes;

            // Starting row (zero‑based index). Row 10 in Excel is index 9.
            int currentRow = 9;
            int columnIndex = 0; // Column A

            // Iterate through all SVG files in the specified folder
            foreach (string svgFile in Directory.GetFiles(svgFolderPath, "*.svg"))
            {
                // Ensure the SVG file exists before attempting to read it
                if (!File.Exists(svgFile))
                {
                    Console.WriteLine($"File not found: {svgFile}");
                    continue;
                }

                try
                {
                    // Open the SVG file as a stream and add it as a picture
                    using (FileStream svgStream = File.OpenRead(svgFile))
                    {
                        shapes.AddPicture(currentRow, columnIndex, currentRow, columnIndex, svgStream);
                    }

                    // Move to the next row for the subsequent icon
                    currentRow++;
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to add picture from '{svgFile}': {ex.Message}");
                }
            }

            // Save the workbook to a file
            string outputPath = "IconsWorkbook.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
