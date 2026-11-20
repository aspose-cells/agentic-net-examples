// Title: C# console utility that loads an Excel workbook and saves it to a new location using Aspose.Cells (without modifying calculation mode)
// Common Searches: how to copy an Excel file to another path with Aspose.Cells in C# | Aspose.Cells .NET load workbook and save without changing calculation settings | C# program to handle missing CalcMode property in Aspose.Cells | save workbook to different folder using Aspose.Cells console app | Aspose.Cells example for loading and saving workbook with error handling
// Tags: Aspose.Cells load workbook C# | Aspose.Cells save workbook to new file | Aspose.Cells calculation mode not exposed .NET | C# console app Excel file copy Aspose.Cells | Aspose.Cells error handling file access

using System;
using System.IO;
using Aspose.Cells;

// The console application accepts input and output file paths, verifies the source file, loads the workbook with Aspose.Cells, ensures the destination directory exists, and saves the workbook unchanged. It includes robust error handling for missing files and runtime exceptions, and notes that the current Aspose.Cells version does not expose a calculation mode property.
class WorkbookUtility
{
    static void Main(string[] args)
    {
        // Ensure input and output file paths are provided
        if (args.Length < 2)
        {
            Console.WriteLine("Usage: WorkbookUtility <inputFilePath> <outputFilePath>");
            return;
        }

        string inputPath = args[0];
        string outputPath = args[1];

        // Verify that the input file exists
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        try
        {
            // Load the workbook from the specified file
            Workbook workbook = new Workbook(inputPath);

            // NOTE: Calculation mode handling removed because the current Aspose.Cells version
            // does not expose the CalcMode property. The workbook is saved unchanged.

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook to the output path
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to: {outputPath}");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors (e.g., file access, Aspose.Cells exceptions)
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
