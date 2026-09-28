// Title: Replace Accent3 theme color with a custom green shade in all XLSX files in a folder using Aspose.Cells for .NET
// AI Prompts: Write C# code that enumerates every .xlsx file in a directory, loads each workbook with Aspose.Cells, changes the Accent3 theme color to RGB(0,150,0) using SetThemeColor, and overwrites the original file. | Show how to apply a custom green shade to the Accent3 theme color across multiple Excel workbooks in a batch process with Aspose.Cells, including error handling for missing or inaccessible files.
// Common Searches: C# batch update Excel theme accent colors with Aspose.Cells | How to programmatically change Accent3 color in many XLSX files | Set custom green shade for Excel theme Accent3 using Aspose.Cells .NET | Iterate over folder of workbooks and modify theme colors in C#
// Tags: setthemecolor accent3 aspocells c# | batch modify excel theme colors aspocells | custom green shade excel theme aspocells | process multiple xlsx files aspocells | replace accent3 theme color c#

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;

// // This program scans a specified folder for .xlsx files, loads each workbook with Aspose.Cells, replaces the Accent3 theme color with a custom green (RGB 0,150,0) via SetThemeColor, and saves the workbook back to its original location.
class Accent3Replacer
{
    static void Main()
    {
        try
        {
            // Folder containing the XLSX files
            string folderPath = @"C:\Path\To\XlsxFolder";

            if (!Directory.Exists(folderPath))
            {
                Console.WriteLine($"Folder not found: {folderPath}");
                return;
            }

            // Define the custom green shade to replace Accent3 (e.g., RGB 0, 150, 0)
            Color customGreen = Color.FromArgb(0, 150, 0);

            // Get all XLSX files in the folder
            string[] files = Directory.GetFiles(folderPath, "*.xlsx", SearchOption.TopDirectoryOnly);

            foreach (string filePath in files)
            {
                try
                {
                    if (!File.Exists(filePath))
                    {
                        Console.WriteLine($"File not found (skipped): {filePath}");
                        continue;
                    }

                    // Load the workbook
                    Workbook workbook = new Workbook(filePath);

                    // Replace Accent3 color in the workbook's theme
                    // Use ThemeColorType enum (available in Aspose.Cells) to specify the theme color
                    workbook.SetThemeColor(ThemeColorType.Accent3, customGreen);

                    // Save the workbook, overwriting the original file
                    workbook.Save(filePath, SaveFormat.Xlsx);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing file '{filePath}': {ex.Message}");
                }
            }

            Console.WriteLine("Accent3 color replacement completed.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}
