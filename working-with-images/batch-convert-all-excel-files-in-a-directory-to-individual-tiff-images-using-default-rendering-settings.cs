// Title: Batch convert every Excel workbook in a folder to separate TIFF images using Aspose.Cells for .NET
// AI Prompts: Write a C# console program that enumerates all .xls, .xlsx, .xlsm, and .xlsb files in a specified directory and saves each workbook as a TIFF file with Aspose.Cells default rendering options. | Enhance the batch conversion script to recursively process subfolders and allow the user to set the output image format (TIFF or PNG) via command‑line arguments.
// Common Searches: C# Aspose.Cells convert multiple Excel files to TIFF in one folder | How to export each worksheet of every workbook in a directory as a TIFF image using Aspose.Cells | Batch image export from Excel to TIFF with default settings in .NET | Automate conversion of all .xlsx files in a folder to .tiff using Aspose.Cells library | Console app to loop through Excel files and save as TIFF images with Aspose.Cells
// Tags: batch Excel to TIFF conversion Aspose.Cells | C# directory processing for workbook image export | save workbook as TIFF default rendering | recursive Excel file conversion to images .NET | command line Aspose.Cells image format selection

using System;
using System.IO;
using Aspose.Cells;

// Demonstrates a C# console application that scans a given input folder, loads each Excel workbook with Aspose.Cells, and saves it as an individual TIFF image in an output folder using the library’s default rendering settings.
class Program
{
    static void Main()
    {
        // Folder containing the source Excel files
        string inputFolder = @"C:\InputExcel";

        // Folder where the TIFF images will be saved
        string outputFolder = @"C:\OutputTiff";

        // Ensure the output directory exists
        Directory.CreateDirectory(outputFolder);

        // Retrieve all files in the input folder
        string[] files = Directory.GetFiles(inputFolder, "*.*", SearchOption.TopDirectoryOnly);

        foreach (string filePath in files)
        {
            // Process only Excel files based on extension
            string ext = Path.GetExtension(filePath).ToLowerInvariant();
            if (ext == ".xls" || ext == ".xlsx" || ext == ".xlsm" || ext == ".xlsb")
            {
                // Load the workbook using Aspose.Cells
                Workbook workbook = new Workbook(filePath);

                // Construct the output TIFF file path (same name, .tiff extension)
                string outputFileName = Path.GetFileNameWithoutExtension(filePath) + ".tiff";
                string outputPath = Path.Combine(outputFolder, outputFileName);

                // Save the workbook as a TIFF image using default rendering settings
                workbook.Save(outputPath, SaveFormat.Tiff);

                Console.WriteLine($"Converted '{filePath}' to '{outputPath}'.");
            }
        }
    }
}
