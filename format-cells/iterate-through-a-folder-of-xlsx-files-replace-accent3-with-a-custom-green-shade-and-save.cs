// Title: Batch replace the Accent3 theme color with a custom green shade across multiple XLSX workbooks using Aspose.Cells for .NET
// Common Searches: C# code to change Accent3 theme color to a specific RGB value in all Excel files in a folder | How to programmatically update Excel theme accent colors for multiple workbooks with Aspose.Cells | Replace Excel Accent3 color with custom green in a batch process using .NET | Iterate through a directory of .xlsx files and modify theme colors with Aspose.Cells API
// Tags: setthemecolor accent3 custom green Aspose.Cells | process multiple xlsx files Aspose.Cells | modify excel theme colors programmatically | iterate folder workbooks .NET | replace theme accent color c#

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;

// The C# utility scans a specified folder for .xlsx files, loads each workbook with Aspose.Cells, uses SetThemeColor to replace the Accent3 theme color with a custom green (RGB 0,176,80), saves the workbook back to the original file, and logs processing results while handling missing files and exceptions.
class Accent3Replacer
{
    static void Main()
    {
        // Folder containing the XLSX files
        string folderPath = @"C:\Path\To\XlsxFolder";

        // Verify the folder exists
        if (!Directory.Exists(folderPath))
        {
            Console.WriteLine($"Folder not found: {folderPath}");
            return;
        }

        // Custom green shade to replace Accent3 (RGB: 0, 176, 80)
        Color customGreen = Color.FromArgb(0, 176, 80);

        // Get all .xlsx files in the folder (non‑recursive)
        string[] files = Directory.GetFiles(folderPath, "*.xlsx");

        foreach (string filePath in files)
        {
            try
            {
                // Ensure the file exists before loading
                if (!File.Exists(filePath))
                {
                    Console.WriteLine($"File not found: {filePath}");
                    continue;
                }

                // Load the workbook
                Workbook workbook = new Workbook(filePath);

                // Replace the Accent3 theme color with the custom green shade
                // Use SetThemeColor to modify the theme accent color
                workbook.SetThemeColor(ThemeColorType.Accent3, customGreen);

                // Save the workbook, overwriting the original file
                workbook.Save(filePath, SaveFormat.Xlsx);

                Console.WriteLine($"Processed: {Path.GetFileName(filePath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing file '{filePath}': {ex.Message}");
            }
        }

        Console.WriteLine("Accent3 color replacement completed.");
    }
}
