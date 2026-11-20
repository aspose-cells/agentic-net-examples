// Title: Batch set workbook language to British English (en‑GB) for all .xlsx files in a folder with Aspose.Cells for .NET
// Common Searches: C# code to change the language of multiple Excel workbooks to en‑GB using Aspose.Cells | How to apply CultureInfo en‑GB to every .xlsx file in a directory with Aspose.Cells | Bulk update Excel file language property in a folder using Aspose.Cells .NET API | Automate setting workbook Settings.CultureInfo for all workbooks in a folder C# | Aspose.Cells example for processing all Excel files in a folder and setting UK English
// Tags: batch set workbook culture Aspose.Cells | update Excel language property en-GB C# | process multiple .xlsx files Aspose.Cells | overwrite workbook after changing CultureInfo | Aspose.Cells bulk document properties editing

using System;
using System.IO;
using System.Globalization;
using Aspose.Cells;

// The sample scans a specified directory for .xlsx files, loads each workbook with Aspose.Cells, assigns the British English culture (en‑GB) via Settings.CultureInfo, saves the workbook back to its original location, and logs successes or errors while handling missing folders or files.
class Program
{
    static void Main(string[] args)
    {
        // Specify the folder containing the workbooks.
        // You can change this path as needed.
        string folderPath = @"C:\Workbooks";

        // Verify that the folder exists to avoid FileNotFoundException.
        if (!Directory.Exists(folderPath))
        {
            Console.WriteLine($"Folder not found: {folderPath}");
            return;
        }

        // Get all Excel files in the folder (you can add other extensions if required).
        string[] excelFiles = Directory.GetFiles(folderPath, "*.xlsx", SearchOption.TopDirectoryOnly);

        foreach (string filePath in excelFiles)
        {
            // Ensure the file still exists before attempting to load.
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"File not found (skipped): {filePath}");
                continue;
            }

            try
            {
                // Load the workbook.
                Workbook workbook = new Workbook(filePath);

                // Set the workbook language to British English using CultureInfo.
                workbook.Settings.CultureInfo = new CultureInfo("en-GB");

                // Save the workbook, overwriting the original file.
                workbook.Save(filePath);
                Console.WriteLine($"Processed: {Path.GetFileName(filePath)}");
            }
            catch (Exception ex)
            {
                // Log any errors for the current file but continue processing others.
                Console.WriteLine($"Error processing '{filePath}': {ex.Message}");
            }
        }

        Console.WriteLine("All workbooks have been processed.");
    }
}
