// Title: Insert SVG icons from a folder into successive rows starting at row 10 using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that reads all *.svg files from a directory and adds each as a picture to column A, beginning at row 10, then saves the workbook with Aspose.Cells. | Generate a method that iterates through a folder of SVG icons, inserts each icon into the next worksheet row starting at index 9, and returns the saved Excel file path.
// Common Searches: Aspose.Cells C# insert multiple SVG images into consecutive rows starting at row 10 | how to batch add SVG icons to Excel worksheet using Aspose.Cells .NET | C# load all SVG files from folder and place them in column A of an Excel file with Aspose.Cells | save workbook after adding pictures from a directory with Aspose.Cells for .NET
// Tags: batch insert svg pictures Aspose.Cells | add svg icons to column A Aspose.Cells | populate worksheet rows with svg images C# | save workbook after picture insertion Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example scans a specified folder for *.svg files, creates a new workbook, and inserts each SVG as a picture into column A of successive rows beginning with row 10, then saves the workbook to the defined output location.
class Program
{
    static void Main()
    {
        try
        {
            // Folder containing SVG icons – change to your actual folder path
            string iconsFolder = @"C:\Icons";

            // Verify the icons folder exists
            if (!Directory.Exists(iconsFolder))
            {
                Console.WriteLine($"Icons folder not found: {iconsFolder}");
                return;
            }

            // Get all SVG files in the folder
            string[] svgFiles = Directory.GetFiles(iconsFolder, "*.svg");

            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Starting row index (row 10 in Excel is index 9 because Aspose.Cells uses zero‑based indexing)
            int startRow = 9;
            int column = 0; // Column A

            // Loop through each SVG file and insert it into successive rows
            for (int i = 0; i < svgFiles.Length; i++)
            {
                string filePath = svgFiles[i];

                // Open the SVG file as a stream and add it as a picture
                using (FileStream stream = new FileStream(filePath, FileMode.Open, FileAccess.Read))
                {
                    sheet.Pictures.Add(startRow + i, column, stream);
                }
            }

            // Ensure the output directory exists
            string outputPath = @"C:\Output\IconsWorkbook.xlsx";
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook to a file
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
