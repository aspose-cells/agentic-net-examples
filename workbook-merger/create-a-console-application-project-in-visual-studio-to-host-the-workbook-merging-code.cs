// Title: C# console program to merge multiple Excel workbooks into one .xlsx file with Aspose.Cells
// AI Prompts: Write a C# console application that accepts an output file path and a list of input .xlsx files as command‑line arguments, loads each workbook with Aspose.Cells, and copies all its worksheets into a new workbook. | Add robust error handling: skip any source file that does not exist, log a warning, create the output directory if it is missing, and catch exceptions during workbook loading and saving. | Use Aspose.Cells.Workbook.Worksheets.AddCopy to duplicate each worksheet by name while preserving the original sheet order, then save the merged workbook in Xlsx format.
// Common Searches: c# console merge several xlsx files using Aspose.Cells library | how to combine multiple Excel workbooks into one workbook in a .NET console app | aspocells copy worksheet by name example for merging workbooks | skip missing Excel files when merging workbooks with Aspose.Cells C# | create output folder automatically before saving merged workbook Aspose.Cells
// Tags: Aspose.Cells merge workbooks C# console | Aspose.Cells AddCopy worksheet duplication | merge multiple .xlsx files programmatically | handle missing source files Aspose.Cells | ensure output directory exists C# Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace WorkbookMergeDemo
{
    // A C# console application that validates command‑line arguments, loads each specified .xlsx workbook with Aspose.Cells, copies all worksheets into a new workbook using the AddCopy method, creates the output folder if needed, and saves the merged workbook while logging warnings for missing files and handling load/save errors.
    class Program
    {
        static void Main(string[] args)
        {
            // Validate input arguments
            if (args.Length < 2)
            {
                Console.WriteLine("Usage: WorkbookMergeDemo <output.xlsx> <input1.xlsx> [<input2.xlsx> ...]");
                return;
            }

            // Output file path
            string outputPath = args[0];

            // Source workbook files (C# 8 range operator)
            string[] sourceFiles = args[1..];

            // Create a new workbook to hold merged content
            Workbook mergedWorkbook = new Workbook();

            // Remove the default empty sheet created by Aspose.Cells
            mergedWorkbook.Worksheets.Clear();

            foreach (string sourceFile in sourceFiles)
            {
                try
                {
                    // Verify source file exists
                    if (!File.Exists(sourceFile))
                    {
                        Console.WriteLine($"Warning: Source file '{sourceFile}' not found. Skipping.");
                        continue;
                    }

                    // Load the source workbook
                    Workbook sourceWorkbook = new Workbook(sourceFile);

                    // Copy each worksheet into the merged workbook
                    foreach (Worksheet sheet in sourceWorkbook.Worksheets)
                    {
                        // AddCopy expects the source sheet name, not the Worksheet object
                        mergedWorkbook.Worksheets.AddCopy(sheet.Name);
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing '{sourceFile}': {ex.Message}");
                }
            }

            try
            {
                // Ensure output directory exists
                string? outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the merged workbook
                mergedWorkbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Merged {sourceFiles.Length} workbook(s) into '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save merged workbook: {ex.Message}");
            }
        }
    }
}
