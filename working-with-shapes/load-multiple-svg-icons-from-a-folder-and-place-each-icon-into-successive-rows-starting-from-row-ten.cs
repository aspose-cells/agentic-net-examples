using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class LoadSvgIcons
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            ShapeCollection shapes = worksheet.Shapes;

            // Path to the folder that contains SVG files
            string folderPath = @"C:\Icons"; // <-- change to your folder

            // Verify that the folder exists
            if (!Directory.Exists(folderPath))
            {
                Console.WriteLine($"Folder not found: {folderPath}");
                return;
            }

            // Retrieve all SVG files in the folder
            string[] svgFiles = Directory.GetFiles(folderPath, "*.svg");

            if (svgFiles.Length == 0)
            {
                Console.WriteLine("No SVG files found in the specified folder.");
                return;
            }

            // Row index where the first icon will be placed (row 10 => zero‑based index 9)
            int startRow = 9;

            // Column index where icons will be placed (e.g., column A => index 0)
            int column = 0;

            // Insert each SVG icon into successive rows
            for (int i = 0; i < svgFiles.Length; i++)
            {
                string filePath = svgFiles[i];

                // Ensure the SVG file exists before reading
                if (!File.Exists(filePath))
                    continue;

                try
                {
                    // Read the SVG file into a byte array
                    byte[] imageData = File.ReadAllBytes(filePath);

                    // Convert byte array to a stream required by AddPicture
                    using (MemoryStream ms = new MemoryStream(imageData))
                    {
                        // Calculate the target row for the current icon
                        int targetRow = startRow + i;

                        // Add the SVG icon to the worksheet.
                        // Height and width set to 0 to use the image's original size.
                        shapes.AddPicture(targetRow, column, 0, 0, ms);
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to insert '{Path.GetFileName(filePath)}': {ex.Message}");
                }
            }

            // Save the workbook with the inserted icons
            string outputPath = "IconsWorkbook.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully: {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}