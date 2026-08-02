// Title: Batch add a disclaimer comment to every worksheet (A1) in multiple Excel files using Aspose.Cells for .NET (C#)
// Description: Scans a given directory for Excel workbooks (.xlsx, .xls, .xlsm, .xlsb), loads each file with Aspose.Cells, iterates through all worksheets, inserts a confidentiality comment into cell A1, saves the workbook in place, and writes processing results to the console.
// Keywords: Aspose.Cells batch comment | C# add Excel comment programmatically | disclaimer note A1 multiple workbooks | iterate worksheets Aspose.Cells | automate Excel disclaimer C#
// Common Searches: how to add the same comment to all sheets in many Excel files using Aspose.Cells | C# batch process Excel workbooks to insert a disclaimer comment | Aspose.Cells loop through worksheets and add notes automatically
// Developer Intent: Insert a standard disclaimer comment into cell A1 of every worksheet across all Excel files in a specified folder.
// Use Cases: Automatically embed a confidentiality notice before distributing a batch of reports. | Enforce compliance by adding a legal disclaimer to every sheet in shared Excel templates. | Prepare financial workbooks with a uniform audit comment in a single automated step.
// AI Prompts: Generate C# code with Aspose.Cells that adds a custom comment to cell A1 of every worksheet in all Excel files within a folder, supporting .xlsx, .xls, .xlsm, and .xlsb extensions. | Explain how to modify the batch script to place the disclaimer in a different cell, set the comment author, or control comment visibility using Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

// Scans a given directory for Excel workbooks (.xlsx, .xls, .xlsm, .xlsb), loads each file with Aspose.Cells, iterates through all worksheets, inserts a confidentiality comment into cell A1, saves the workbook in place, and writes processing results to the console.
class BatchAddDisclaimer
{
    // Standard disclaimer text to be added as a comment
    private const string Disclaimer = "Confidential: This workbook contains proprietary information.";

    static void Main()
    {
        // Folder containing the Excel workbooks to process
        string inputFolder = @"C:\InputWorkbooks";

        // Validate folder existence
        if (!Directory.Exists(inputFolder))
        {
            Console.WriteLine($"Folder not found: {inputFolder}");
            return;
        }

        // Process each Excel file in the folder (supports common extensions)
        foreach (string filePath in Directory.GetFiles(inputFolder, "*.*", SearchOption.TopDirectoryOnly))
        {
            string ext = Path.GetExtension(filePath).ToLowerInvariant();
            if (ext != ".xlsx" && ext != ".xls" && ext != ".xlsm" && ext != ".xlsb")
                continue; // Skip non‑Excel files

            try
            {
                // Load the workbook using the provided constructor
                using (Workbook workbook = new Workbook(filePath))
                {
                    // Iterate through all worksheets
                    foreach (Worksheet sheet in workbook.Worksheets)
                    {
                        // Add a comment to the top‑left cell (A1)
                        int commentIndex = sheet.Comments.Add("A1");
                        sheet.Comments[commentIndex].Note = Disclaimer;
                    }

                    // Save the modified workbook back to the same file using the provided Save method
                    workbook.Save(filePath);
                }

                Console.WriteLine($"Processed: {Path.GetFileName(filePath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing {Path.GetFileName(filePath)}: {ex.Message}");
            }
        }

        Console.WriteLine("Batch processing completed.");
    }
}
