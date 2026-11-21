// Title: Batch remove hidden worksheets from Excel workbooks with Aspose.Cells C# only when a confidential sheet is present
// Common Searches: C# Aspose.Cells batch delete hidden sheets if workbook contains confidential worksheet | How to process multiple .xlsx files and remove invisible worksheets conditionally using Aspose.Cells | Aspose.Cells remove hidden worksheets during bulk export when sheet name includes confidential | Iterate through folder of Excel files in C# and delete hidden tabs only for confidential workbooks | Conditional hidden worksheet removal with Aspose.Cells before saving workbook
// Tags: batch hidden worksheet removal Aspose.Cells | conditional workbook processing confidential sheet | iterate Excel files folder C# Aspose.Cells | remove invisible worksheets before save | detect worksheet name contains confidential Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace WorkbookBatchProcessor
{
    // The C# program scans a source folder for .xlsx files, loads each workbook with Aspose.Cells, checks for any worksheet whose name includes "confidential" (case‑insensitive), and if found, removes all hidden worksheets before saving the workbook to a destination directory.
    class Program
    {
        static void Main(string[] args)
        {
            // Define source and destination directories
            string sourceDirectory = @"C:\InputWorkbooks";
            string destinationDirectory = @"C:\OutputWorkbooks";

            // Verify source directory exists
            if (!Directory.Exists(sourceDirectory))
            {
                Console.WriteLine($"Source directory does not exist: {sourceDirectory}");
                return;
            }

            // Ensure the destination directory exists
            if (!Directory.Exists(destinationDirectory))
            {
                Directory.CreateDirectory(destinationDirectory);
            }

            // Process each Excel file in the source directory
            foreach (string filePath in Directory.GetFiles(sourceDirectory, "*.xlsx"))
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

                    // Determine if the workbook contains any worksheet with "confidential" in its name
                    bool containsConfidential = false;
                    foreach (Worksheet sheet in workbook.Worksheets)
                    {
                        if (sheet.Name.IndexOf("confidential", StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            containsConfidential = true;
                            break;
                        }
                    }

                    // If confidential, remove hidden worksheets before saving
                    if (containsConfidential)
                    {
                        try
                        {
                            for (int i = workbook.Worksheets.Count - 1; i >= 0; i--)
                            {
                                if (!workbook.Worksheets[i].IsVisible)
                                {
                                    workbook.Worksheets.RemoveAt(i);
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"Error removing hidden worksheets in '{Path.GetFileName(filePath)}': {ex.Message}");
                        }
                    }

                    // Prepare the destination path
                    string destinationPath = Path.Combine(destinationDirectory, Path.GetFileName(filePath));

                    // Save the workbook
                    workbook.Save(destinationPath);

                    Console.WriteLine($"Processed: {Path.GetFileName(filePath)}");
                }
                catch (Exception ex)
                {
                    // Log any errors for the current file and continue processing others
                    Console.WriteLine($"Error processing file '{filePath}': {ex.Message}");
                }
            }
        }
    }
}
