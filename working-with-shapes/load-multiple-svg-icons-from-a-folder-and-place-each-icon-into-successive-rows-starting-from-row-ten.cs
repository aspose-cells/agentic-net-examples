// Title: Load SVG icons from a folder and insert each into successive rows starting at row 10 with Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells insert SVG files from a directory into specific worksheet rows | How to add multiple SVG pictures to Excel starting at row 10 using Aspose.Cells | Aspose.Cells load all SVG icons in a folder and place them sequentially in column A | Resize SVG picture to match row height when inserting with Aspose.Cells C# | Save workbook after inserting SVG icons with Aspose.Cells .NET
// Tags: load svg files from folder Aspose.Cells | add svg picture to worksheet Aspose.Cells | insert pictures into successive rows Excel C# | resize svg picture to row height Aspose.Cells | save workbook as xlsx Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace SvgIconInserter
{
    // The example scans a specified folder for SVG files, creates a new workbook, and inserts each SVG into column A of successive rows beginning with Excel row 10. It optionally resizes each picture to the default row height, writes the file name in the adjacent cell, and saves the workbook as an XLSX file using Aspose.Cells for .NET.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the folder that contains SVG files.
            // Change this path to point to your actual folder.
            string svgFolderPath = Path.Combine(Environment.CurrentDirectory, "Icons");

            if (!Directory.Exists(svgFolderPath))
            {
                Console.WriteLine($"Error: The folder \"{svgFolderPath}\" does not exist.");
                return;
            }

            // Get all SVG files in the folder (non‑recursive).
            string[] svgFiles = Directory.GetFiles(svgFolderPath, "*.svg", SearchOption.TopDirectoryOnly);
            if (svgFiles.Length == 0)
            {
                Console.WriteLine($"No SVG files found in \"{svgFolderPath}\".");
                return;
            }

            // Create a new workbook and get the first worksheet.
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Starting row (zero‑based). Row ten in Excel corresponds to index 9.
            int currentRow = 9;
            int columnIndex = 0; // Column A

            foreach (string svgPath in svgFiles)
            {
                try
                {
                    using (FileStream svgStream = new FileStream(svgPath, FileMode.Open, FileAccess.Read))
                    {
                        // Add the SVG picture to the worksheet.
                        // The Add method returns the index of the picture in the Pictures collection.
                        int pictureIndex = sheet.Pictures.Add(currentRow, columnIndex, svgStream);
                        Picture picture = sheet.Pictures[pictureIndex];

                        // Optional: adjust picture size to fit the cell height.
                        // Here we set the picture height to the default row height (15 points) and keep aspect ratio.
                        double defaultRowHeight = sheet.Cells.GetRowHeightPixel(currentRow);
                        picture.Height = (int)defaultRowHeight;
                        picture.Width = picture.Height; // simple square scaling; adjust as needed.

                        // Optionally, write the file name next to the picture for reference.
                        sheet.Cells[currentRow, columnIndex + 1].PutValue(Path.GetFileName(svgPath));
                    }

                    Console.WriteLine($"Inserted \"{Path.GetFileName(svgPath)}\" at row {currentRow + 1}.");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to insert \"{Path.GetFileName(svgPath)}\": {ex.Message}");
                }

                // Move to the next row for the next icon.
                currentRow++;
            }

            // Save the workbook.
            string outputPath = Path.Combine(Environment.CurrentDirectory, "SvgIcons.xlsx");
            try
            {
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save workbook: {ex.Message}");
            }
        }
    }
}
