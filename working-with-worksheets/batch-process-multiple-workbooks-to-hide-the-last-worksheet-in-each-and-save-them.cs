// Title: Hide the last worksheet in each Excel workbook of a folder using Aspose.Cells for .NET
// AI Prompts: Write a C# console program that scans a specified input directory for *.xlsx files, loads each workbook with Aspose.Cells, sets the IsVisible property of the last worksheet to false, and saves the modified file to an output directory while preserving the original filename. | Extend the batch routine to hide the last worksheet only when its name starts with "Temp" and output the processed file names to the console.
// Common Searches: Aspose.Cells hide last sheet in multiple workbooks C# | C# batch process Excel files to change worksheet visibility with Aspose | How to programmatically set worksheet IsVisible false for all files in a folder using Aspose.Cells | Iterate over .xlsx files in a directory and hide specific worksheets in .NET
// Tags: batch hide last worksheet Aspose.Cells | process multiple .xlsx files C# | set worksheet IsVisible false Aspose.Cells | save modified workbooks to output folder | conditional hide worksheet by name Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// // Loads each .xlsx file from an input folder, hides its last worksheet by setting IsVisible = false, and saves the updated workbook to a designated output folder while keeping the original file name.
class BatchHideLastWorksheet
{
    static void Main()
    {
        // Folder containing the workbooks to process
        string inputFolder = @"C:\Workbooks\Input";
        // Folder where the processed workbooks will be saved
        string outputFolder = @"C:\Workbooks\Output";

        // Ensure output folder exists
        Directory.CreateDirectory(outputFolder);

        // Get all Excel files in the input folder (you can adjust the pattern as needed)
        string[] files = Directory.GetFiles(inputFolder, "*.xlsx", SearchOption.TopDirectoryOnly);

        foreach (string filePath in files)
        {
            // Load the workbook
            Workbook workbook = new Workbook(filePath);

            // Get the last worksheet (index is Count - 1 because collection is zero‑based)
            int lastIndex = workbook.Worksheets.Count - 1;
            if (lastIndex >= 0)
            {
                // Hide the last worksheet
                Worksheet lastSheet = workbook.Worksheets[lastIndex];
                lastSheet.IsVisible = false;
            }

            // Build the output file path (preserve original file name)
            string fileName = Path.GetFileName(filePath);
            string outputPath = Path.Combine(outputFolder, fileName);

            // Save the modified workbook
            workbook.Save(outputPath);
        }

        Console.WriteLine("Processing completed. {0} files updated.", files.Length);
    }
}
