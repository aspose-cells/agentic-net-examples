// Title: Batch update Excel workbook language to British English with Aspose.Cells for .NET
// AI Prompts: Write a C# console program that enumerates all .xls, .xlsx, .xlsm, and .xlsb files in a given directory, loads each workbook with Aspose.Cells, sets workbook.Settings.CultureInfo to "en-GB", and saves the file in place. | Show how to handle missing files and unsupported extensions while applying a British English locale to multiple Excel workbooks using Aspose.Cells in a loop.
// Common Searches: how to change the language of many Excel files to en-GB using Aspose.Cells in C# | C# code to set workbook CultureInfo for all Excel workbooks in a folder | batch process Excel workbooks with Aspose.Cells to apply British English locale | Aspose.Cells loop through directory and update workbook language property | overwrite original Excel files after setting CultureInfo with Aspose.Cells .NET
// Tags: set workbook CultureInfo Aspose.Cells C# | batch process Excel files Aspose.Cells .NET | apply British English locale to Excel workbooks | overwrite original workbook after property change | iterate folder Excel files Aspose.Cells

using System;
using System.IO;
using System.Globalization;
using Aspose.Cells;

// The C# console app scans a specified folder, loads each supported Excel workbook (.xls, .xlsx, .xlsm, .xlsb) with Aspose.Cells, sets its Settings.CultureInfo to British English (en-GB), saves the workbook back to the original file, and logs success or error messages for each file.
class Program
{
    static void Main()
    {
        // Folder containing the workbooks to process
        string folderPath = @"C:\Path\To\Folder"; // TODO: set your folder path

        if (!Directory.Exists(folderPath))
        {
            Console.WriteLine($"Folder not found: {folderPath}");
            return;
        }

        // Retrieve all Excel files in the folder
        string[] files = Directory.GetFiles(folderPath, "*.*", SearchOption.TopDirectoryOnly);

        foreach (string file in files)
        {
            // Process only supported Excel extensions
            string ext = Path.GetExtension(file).ToLowerInvariant();
            if (ext != ".xls" && ext != ".xlsx" && ext != ".xlsm" && ext != ".xlsb")
                continue;

            if (!File.Exists(file))
            {
                Console.WriteLine($"File not found (skipped): {file}");
                continue;
            }

            try
            {
                // Load the workbook
                Workbook workbook = new Workbook(file);

                // Set the culture to British English
                workbook.Settings.CultureInfo = new CultureInfo("en-GB");

                // Save the workbook, overwriting the original file
                workbook.Save(file);
                Console.WriteLine($"Processed: {file}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing file '{file}': {ex.Message}");
            }
        }
    }
}
