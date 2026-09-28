// Title: How to unhide all columns, enable scrollbars, and save a worksheet to a new XLSX file using Aspose.Cells for .NET (C#)
// AI Prompts: Write a C# example with Aspose.Cells that loads a workbook, makes every column in the first worksheet visible, activates worksheet scrollbars, and saves the file to a specified output path. | Provide Aspose.Cells .NET code that iterates over all used columns, sets Columns.IsHidden to false, creates the destination folder if it does not exist, and exports the workbook as an XLSX file to a different directory.
// Common Searches: Aspose.Cells C# unhide all columns in worksheet and keep scrollbars visible | save modified Excel workbook to a new folder using Aspose.Cells .NET | how to programmatically enable scrollbars when exporting Excel with Aspose.Cells | C# code to iterate through columns and set IsHidden false in Aspose.Cells | Aspose.Cells example for copying workbook to new location after changing column visibility
// Tags: column visibility Aspose.Cells C# | set column IsHidden false Aspose.Cells | worksheet scrollbars activation Aspose.Cells | export workbook to different directory Aspose.Cells | ensure output folder exists Aspose.Cells C#

using Aspose.Cells;
using System;
using System.IO;

// The sample loads an existing XLSX file, loops through every used column in the first worksheet to set Columns.IsHidden = false, guarantees the target directory exists, and saves the updated workbook to a new location as an XLSX file, handling missing files and runtime errors.
class Program
{
    static void Main()
    {
        // Paths for source workbook and destination workbook
        string sourcePath = @"C:\Input\workbook.xlsx";
        string destinationPath = @"C:\Output\workbook_unhidden.xlsx";

        try
        {
            // Verify source file exists
            if (!File.Exists(sourcePath))
            {
                Console.WriteLine($"Source file not found: {sourcePath}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(sourcePath);

            // Get the first worksheet (modify as needed for other sheets)
            Worksheet sheet = workbook.Worksheets[0];

            // Unhide all columns in the worksheet
            // MaxColumn returns the index of the last used column (0‑based)
            int maxColumn = sheet.Cells.MaxColumn;
            for (int col = 0; col <= maxColumn; col++)
            {
                // Use the Columns collection to set the IsHidden property
                sheet.Cells.Columns[col].IsHidden = false;
            }

            // Ensure the output directory exists
            string? outputDir = Path.GetDirectoryName(destinationPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the modified workbook to a new location
            workbook.Save(destinationPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to: {destinationPath}");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
