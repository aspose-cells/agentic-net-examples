// Title: Create a C# console application using Aspose.Cells to merge multiple .xlsx workbooks into a single file
// Common Searches: how to build a .NET console program that merges several Excel workbooks with Aspose.Cells | Aspose.Cells C# command line utility to combine multiple .xlsx files into one workbook | copy worksheets from multiple source workbooks into a new workbook using Aspose.Cells API | C# program to merge Excel files and ensure the output folder is created automatically | using Aspose.Cells AddCopy method to merge worksheets from different workbooks
// Tags: Aspose.Cells merge workbooks console app | C# Aspose.Cells AddCopy worksheets | command line Excel workbook merger .NET | create output directory Aspose.Cells save workbook | validate command line arguments C# Excel merge

using System;
using System.IO;
using Aspose.Cells;

namespace WorkbookMergingApp
{
    // A console application that validates command‑line arguments, creates an empty Workbook, clears the default sheet, loads each input .xlsx file, copies every worksheet into the new workbook using AddCopy, ensures the output directory exists, and saves the merged workbook to the specified path.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Validate input arguments
                if (args.Length < 3)
                {
                    Console.WriteLine("Usage: WorkbookMergingApp <output.xlsx> <input1.xlsx> <input2.xlsx> [<input3.xlsx> ...]");
                    return;
                }

                // First argument is the output workbook path
                string outputPath = args[0];

                // Remaining arguments are input workbooks to merge
                string[] inputPaths = new string[args.Length - 1];
                Array.Copy(args, 1, inputPaths, 0, inputPaths.Length);

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(outputPath) ?? string.Empty;
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Create a new empty workbook that will hold the merged result
                Workbook mergedWorkbook = new Workbook();

                // Remove the default empty worksheet created by Aspose.Cells
                mergedWorkbook.Worksheets.Clear();

                // Iterate over each input workbook
                foreach (string inputPath in inputPaths)
                {
                    if (!File.Exists(inputPath))
                    {
                        Console.WriteLine($"Input file not found: {inputPath}");
                        continue;
                    }

                    try
                    {
                        // Load the source workbook
                        Workbook sourceWorkbook = new Workbook(inputPath);

                        // Copy each worksheet from the source workbook into the merged workbook
                        foreach (Worksheet sheet in sourceWorkbook.Worksheets)
                        {
                            // AddCopy expects the source sheet name (string) in this API version
                            mergedWorkbook.Worksheets.AddCopy(sheet.Name);
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Failed to process '{inputPath}': {ex.Message}");
                    }
                }

                // Save the merged workbook to the specified output path
                mergedWorkbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Merged workbook saved to: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
