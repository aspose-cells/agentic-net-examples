using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        try
        {
            // Folder containing the source XLSX files
            string sourceFolder = @"C:\SourceFolder";

            // Folder where the modified files will be saved
            string destinationFolder = @"C:\DestinationFolder";

            // Verify source folder exists
            if (!Directory.Exists(sourceFolder))
            {
                Console.WriteLine($"Source folder not found: {sourceFolder}");
                return;
            }

            // Ensure the destination folder exists
            Directory.CreateDirectory(destinationFolder);

            // Define the custom green shade to replace Accent3
            Color customGreen = Color.FromArgb(0, 150, 0); // adjust RGB as needed

            // Process each .xlsx file in the source folder
            foreach (string filePath in Directory.GetFiles(sourceFolder, "*.xlsx"))
            {
                try
                {
                    // Verify the file exists before loading
                    if (!File.Exists(filePath))
                    {
                        Console.WriteLine($"File not found: {filePath}");
                        continue;
                    }

                    // Load the workbook
                    Workbook workbook = new Workbook(filePath);

                    // Replace the Accent3 theme color with the custom green
                    workbook.SetThemeColor(ThemeColorType.Accent3, customGreen);

                    // Determine the output file path
                    string fileName = Path.GetFileName(filePath);
                    string outputPath = Path.Combine(destinationFolder, fileName);

                    // Save the modified workbook
                    workbook.Save(outputPath, SaveFormat.Xlsx);
                }
                catch (Exception exFile)
                {
                    Console.WriteLine($"Error processing file '{filePath}': {exFile.Message}");
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}