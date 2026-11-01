// Title: C# console program to batch add a custom document property to multiple Excel .xlsx workbooks and output a success/failure summary using Aspose.Cells
// Common Searches: C# code to add a custom document property to all .xlsx files in a folder using Aspose.Cells | how to batch modify Excel workbooks and log successes and failures with Aspose.Cells | Aspose.Cells example for updating custom properties across multiple workbooks and creating a summary report | process a directory of Excel files in C# and add a standard library reference property | automate adding library reference to many workbooks and get a processing summary using Aspose.Cells
// Tags: custom document property insertion Aspose.Cells | batch workbook metadata update Aspose.Cells | process multiple xlsx files Aspose.Cells | generate processing summary console application | save updated workbook to separate folder Aspose.Cells

using System;
using System.IO;
using System.Collections.Generic;
using Aspose.Cells;

namespace BatchWorkbookProcessor
{
    // The program scans a specified input folder for .xlsx files, loads each workbook with Aspose.Cells, inserts or updates a custom document property named 'StandardLibrary' with the value 'MyStandardLib', saves the modified workbook to an output folder, and prints a concise summary reporting total files processed, successful updates, and any failures.
    class Program
    {
        // Path to the folder containing workbooks to process
        private const string InputFolder = @"C:\Workbooks\Input";
        // Path to the folder where processed workbooks will be saved (can be same as input)
        private const string OutputFolder = @"C:\Workbooks\Output";

        // Standard library reference to add to each workbook
        private const string LibraryReferenceName = "StandardLibrary";
        private const string LibraryReferenceValue = "MyStandardLib";

        static void Main()
        {
            try
            {
                // Ensure input folder exists
                if (!Directory.Exists(InputFolder))
                {
                    Console.WriteLine($"Input folder does not exist: {InputFolder}");
                    return;
                }

                // Ensure output folder exists
                Directory.CreateDirectory(OutputFolder);

                // Collect all Excel files in the input folder (top‑level only)
                string[] workbookFiles = Directory.GetFiles(InputFolder, "*.xlsx", SearchOption.TopDirectoryOnly);

                // Lists to keep track of processing results
                List<string> successList = new List<string>();
                List<string> failureList = new List<string>();

                foreach (string inputPath in workbookFiles)
                {
                    try
                    {
                        // Verify the file still exists before loading
                        if (!File.Exists(inputPath))
                        {
                            failureList.Add($"{Path.GetFileName(inputPath)} : File not found");
                            continue;
                        }

                        // Load the workbook (lifecycle rule: load)
                        Workbook workbook = new Workbook(inputPath);

                        // Add or update the standard library reference as a custom document property
                        // (feature: add library reference)
                        var customProps = workbook.CustomDocumentProperties;
                        var existingProp = customProps[LibraryReferenceName];
                        if (existingProp != null)
                        {
                            existingProp.Value = LibraryReferenceValue;
                        }
                        else
                        {
                            customProps.Add(LibraryReferenceName, LibraryReferenceValue);
                        }

                        // Determine output path (same file name, different folder)
                        string outputPath = Path.Combine(OutputFolder, Path.GetFileName(inputPath));

                        // Save the modified workbook (lifecycle rule: save)
                        workbook.Save(outputPath, SaveFormat.Xlsx);

                        successList.Add(Path.GetFileName(inputPath));
                    }
                    catch (Exception ex)
                    {
                        // Record failure with exception message for diagnostics
                        failureList.Add($"{Path.GetFileName(inputPath)} : {ex.Message}");
                    }
                }

                // Generate summary report
                Console.WriteLine("Batch Processing Summary");
                Console.WriteLine("------------------------");
                Console.WriteLine($"Total files processed : {workbookFiles.Length}");
                Console.WriteLine($"Successful updates    : {successList.Count}");
                Console.WriteLine($"Failed updates        : {failureList.Count}");
                Console.WriteLine();

                if (successList.Count > 0)
                {
                    Console.WriteLine("Successfully processed files:");
                    foreach (var file in successList)
                    {
                        Console.WriteLine($" - {file}");
                    }
                    Console.WriteLine();
                }

                if (failureList.Count > 0)
                {
                    Console.WriteLine("Files that failed to process:");
                    foreach (var entry in failureList)
                    {
                        Console.WriteLine($" - {entry}");
                    }
                }
            }
            catch (Exception ex)
            {
                // Top‑level exception handling
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }
    }
}
