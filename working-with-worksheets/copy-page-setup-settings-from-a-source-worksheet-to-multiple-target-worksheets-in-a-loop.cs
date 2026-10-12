// Title: Copy page setup from a source worksheet to multiple target worksheets in a loop using Aspose.Cells for .NET (C#)
// AI Prompts: Load a workbook, retrieve the print settings (PageSetup) of a worksheet named 'Source', and apply them to a list of target worksheets, creating any that are missing, with Aspose.Cells in C#. | Use Aspose.Cells CopyOptions to replicate the worksheet's print settings (orientation, margins, headers/footers) from one sheet to several others within the same workbook. | After copying the print settings to all target sheets, save the workbook to a new file path using Aspose.Cells.
// Common Searches: Aspose.Cells copy page setup from one sheet to many sheets in C# | C# loop to apply same page setup to multiple worksheets using Aspose.Cells | How to add missing worksheets and duplicate page setup with Aspose.Cells .NET
// Tags: Aspose.Cells worksheet print settings replication | C# batch copy worksheet print configuration | worksheet print settings copy with Aspose.Cells | add missing worksheets programmatically Aspose.Cells | save modified workbook Aspose.Cells .NET

using System;
using System.IO;
using Aspose.Cells;

// The program loads 'source.xlsx', checks for a worksheet named 'Source', iterates over an array of target sheet names, adds any missing worksheets, copies the source worksheet's PageSetup to each target using default CopyOptions, and saves the updated workbook as 'output.xlsx'.
class Program
{
    static void Main()
    {
        try
        {
            // Path to the source workbook
            string sourcePath = "source.xlsx";

            // Verify that the source file exists to avoid FileNotFoundException
            if (!File.Exists(sourcePath))
            {
                Console.WriteLine($"Error: The source file '{sourcePath}' was not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(sourcePath);

            // Source worksheet whose page setup will be copied
            Worksheet sourceSheet = workbook.Worksheets["Source"];
            if (sourceSheet == null)
            {
                Console.WriteLine("Error: Worksheet named 'Source' does not exist in the workbook.");
                return;
            }

            // Names of target worksheets
            string[] targetSheetNames = { "Target1", "Target2", "Target3" };

            // Options for copying page setup (default settings)
            CopyOptions copyOptions = new CopyOptions();

            foreach (string sheetName in targetSheetNames)
            {
                // Get the target worksheet if it exists; otherwise, add a new one
                Worksheet targetSheet = workbook.Worksheets[sheetName];
                if (targetSheet == null)
                {
                    targetSheet = workbook.Worksheets.Add(sheetName);
                }

                // Copy page setup settings from source to target
                targetSheet.PageSetup.Copy(sourceSheet.PageSetup, copyOptions);
            }

            // Save the workbook
            string outputPath = "output.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Handle unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
