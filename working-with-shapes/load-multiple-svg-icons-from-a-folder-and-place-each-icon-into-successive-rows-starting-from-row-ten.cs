// Title: Load SVG icons from a folder and insert each into successive rows starting at row 10 using Aspose.Cells for .NET
// AI Prompts: Write C# code with Aspose.Cells that enumerates all *.svg files in a directory and adds each as a picture to column A, beginning at Excel row 10 and moving down one row per icon. | Include logic to create the target output folder if it does not exist and save the workbook as an XLSX file to a given path.
// Common Searches: aspnet c# insert multiple svg images into Excel rows starting at row 10 | how to batch add pictures from a folder to an Aspose.Cells worksheet | Aspose.Cells load all svg files from directory into column A | C# save workbook after adding images and ensure output directory exists | place icons in successive rows in Excel using Aspose.Cells API
// Tags: add svg pictures to worksheet Aspose.Cells | batch insert images into Excel rows C# | create output directory before saving workbook | populate column A with icons starting at row 10 | enumerate files from folder Aspose.Cells picture insertion

using Aspose.Cells;
using System;
using System.IO;

// The example scans a specified folder for SVG files, creates a new workbook, and inserts each SVG as a picture into column A starting at Excel row 10, advancing one row per icon. It also ensures the output directory exists before saving the workbook as an XLSX file.
class Program
{
    static void Main()
    {
        try
        {
            // Folder containing the SVG icons
            string folderPath = @"C:\Icons";

            // Verify the source folder exists
            if (!Directory.Exists(folderPath))
            {
                Console.WriteLine($"Source folder not found: {folderPath}");
                return;
            }

            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Excel rows are zero‑based in Aspose.Cells.
            // Row 10 in Excel corresponds to index 9.
            int currentRow = 9;
            int columnIndex = 0; // Column A

            // Retrieve all SVG files from the folder
            string[] svgFiles = Directory.GetFiles(folderPath, "*.svg");

            foreach (string svgPath in svgFiles)
            {
                // Open the SVG file as a stream and insert it into the worksheet
                using (FileStream stream = new FileStream(svgPath, FileMode.Open, FileAccess.Read))
                {
                    sheet.Pictures.Add(currentRow, columnIndex, stream);
                }

                // Advance to the next row for the subsequent icon
                currentRow++;
            }

            // Ensure the output directory exists
            string outputPath = @"C:\Output\IconsWorkbook.xlsx";
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook to a file
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
