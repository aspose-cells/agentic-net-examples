// Title: Batch Insert SVG Icons into Excel Rows Starting at Row 10 with Aspose.Cells for .NET (C#)
// Description: This C# example scans a folder for *.svg files, reads each icon into a memory stream, and adds it as a picture shape to column A of the first worksheet. Icons are placed in consecutive rows beginning with row 10 (zero‑based index 9). The workbook is then saved as an Excel file.
// Keywords: Aspose.Cells | C# | .NET | SVG insertion | batch add images | Excel picture shape | load SVG from folder | place icons in rows | Excel automation | shape placement
// Common Searches: Aspose.Cells add multiple SVG files to Excel | C# insert SVG icons into consecutive rows | batch load SVG images into worksheet using Aspose.Cells | place picture shapes starting at row 10 in Excel | how to embed SVG icons in Excel with .NET
// Developer Intent: Read all SVG files from a directory and embed each one as a picture in successive rows beginning at row 10.
// Use Cases: Generate an icon catalog where each row from 10 onward shows a different SVG thumbnail in column A. | Create a product sheet with SVG previews aligned vertically, starting at row 10. | Automate a style‑guide workbook by inserting SVG symbols sequentially into cells.
// AI Prompts: Write C# code using Aspose.Cells to load every SVG file from a given folder and insert each into column A, one per row starting at row 10, with error handling for missing files. | Show how to automatically resize each inserted SVG shape to fit the cell height while preserving its aspect ratio. | Explain how to modify the sample to place icons in column B and start at row 20 instead of column A and row 10.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// This C# example scans a folder for *.svg files, reads each icon into a memory stream, and adds it as a picture shape to column A of the first worksheet. Icons are placed in consecutive rows beginning with row 10 (zero‑based index 9). The workbook is then saved as an Excel file.
class LoadSvgIcons
{
    static void Main()
    {
        try
        {
            // Path to the folder containing SVG files – change to a valid folder on your machine
            string folderPath = @"C:\Icons";

            // Verify that the folder exists
            if (!Directory.Exists(folderPath))
            {
                Console.WriteLine($"Folder not found: {folderPath}");
                return;
            }

            // Retrieve all SVG files from the folder
            string[] svgFiles = Directory.GetFiles(folderPath, "*.svg");
            if (svgFiles.Length == 0)
            {
                Console.WriteLine("No SVG files found in the specified folder.");
                return;
            }

            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            ShapeCollection shapes = worksheet.Shapes;

            // Starting row (zero‑based index). Row 10 => index 9
            int startRow = 9;
            // Column where icons will be placed (A => index 0)
            int column = 0;

            // Insert each SVG icon into successive rows
            for (int i = 0; i < svgFiles.Length; i++)
            {
                // Ensure the SVG file exists before reading
                if (!File.Exists(svgFiles[i]))
                {
                    Console.WriteLine($"File not found: {svgFiles[i]}");
                    continue;
                }

                // Read the SVG file into a byte array
                byte[] imageData = File.ReadAllBytes(svgFiles[i]);

                // Add the SVG icon to the worksheet using a memory stream
                using (MemoryStream ms = new MemoryStream(imageData))
                {
                    // Define a small area for the picture; Excel will scale it automatically
                    int upperLeftRow = startRow + i;
                    int upperLeftColumn = column;
                    int lowerRightRow = upperLeftRow + 5;   // arbitrary height
                    int lowerRightColumn = upperLeftColumn + 5; // arbitrary width

                    Shape shape = shapes.AddPicture(upperLeftRow, upperLeftColumn, lowerRightRow, lowerRightColumn, ms);
                    // Optional: let Excel auto‑size the shape
                    shape.Placement = PlacementType.FreeFloating;
                }
            }

            // Save the workbook with the inserted icons
            string outputPath = "IconsWorkbook.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {Path.GetFullPath(outputPath)}");
        }
        catch (DirectoryNotFoundException dirEx)
        {
            Console.WriteLine($"Directory error: {dirEx.Message}");
        }
        catch (FileNotFoundException fileEx)
        {
            Console.WriteLine($"File error: {fileEx.Message}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}
