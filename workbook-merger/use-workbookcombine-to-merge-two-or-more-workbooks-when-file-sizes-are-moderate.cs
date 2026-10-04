// Title: Merge several Excel workbooks into a single .xlsx file using Aspose.Cells Workbook.Combine in C#
// AI Prompts: Generate a C# console program that reads an array of .xlsx file paths, verifies each file exists, loads the available workbooks with Aspose.Cells, merges them using Workbook.Combine, and saves the result as Combined.xlsx. | Create a C# snippet that shows how to combine multiple Excel workbooks with Aspose.Cells while gracefully skipping missing files and writing the merged workbook to an Xlsx document.
// Common Searches: c# Aspose.Cells combine multiple workbooks into one file | how to skip missing Excel files when merging with Aspose.Cells | combine three .xlsx files using Aspose.Cells Workbook.Combine | sample code for merging Excel workbooks in C# with Aspose.Cells | Aspose.Cells merge workbooks moderate file size
// Tags: Aspose.Cells Workbook.Combine for merging workbooks | C# merge multiple Excel files using Aspose.Cells | handle absent source files during workbook combine | output combined workbook as Xlsx format | moderate size workbook consolidation in C#

using Aspose.Cells;
using System;
using System.IO;

// The program checks each specified .xlsx path, loads existing workbooks with Aspose.Cells, uses the first workbook as the base and merges subsequent ones via Workbook.Combine, skips files that are not found, and saves the consolidated workbook as Combined.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Paths of the workbooks to be merged
            string[] sourceFiles = { "Book1.xlsx", "Book2.xlsx", "Book3.xlsx" };

            // Verify that at least one source file exists
            bool anyFileExists = false;
            foreach (string path in sourceFiles)
            {
                if (File.Exists(path))
                {
                    anyFileExists = true;
                    break;
                }
            }

            if (!anyFileExists)
            {
                Console.WriteLine("No source files were found. Operation aborted.");
                return;
            }

            Workbook combinedWorkbook = null;
            bool isFirst = true;

            // Load each existing workbook and combine them
            foreach (string filePath in sourceFiles)
            {
                if (!File.Exists(filePath))
                {
                    Console.WriteLine($"File not found and will be skipped: {filePath}");
                    continue;
                }

                Workbook wb = new Workbook(filePath);

                if (isFirst)
                {
                    // Use the first workbook as the base
                    combinedWorkbook = wb;
                    isFirst = false;
                }
                else
                {
                    // Merge subsequent workbooks into the base workbook
                    combinedWorkbook.Combine(wb);
                }
            }

            // If no workbook was loaded, create an empty one
            if (combinedWorkbook == null)
            {
                combinedWorkbook = new Workbook();
            }

            // Save the combined workbook to a new file
            string outputPath = "Combined.xlsx";
            combinedWorkbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Combined workbook saved to: {outputPath}");
        }
        catch (Exception ex)
        {
            // Handle unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
