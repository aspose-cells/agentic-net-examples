// Title: C# – Batch replace Accent3 theme color with a custom green shade in multiple XLSX files using Aspose.Cells
// Description: Iterates through a specified folder, loads each .xlsx workbook with Aspose.Cells, changes the Accent3 theme color to a user‑defined green (RGB 0,128,0) via SetThemeColor, saves the file, and logs results while handling missing files and exceptions.
// Keywords: Aspose.Cells | C# | SetThemeColor | Accent3 | theme color | batch update Excel | XLSX folder processing | custom green shade | RGB 0 128 0 | automate Excel theme | replace theme color programmatically
// Common Searches: change Accent3 theme color in all Excel files using Aspose.Cells | C# code to update Excel theme colors in a folder | batch modify XLSX theme colors with Aspose.Cells .NET | set custom green shade for Accent3 across multiple workbooks | automate Excel theme color replacement C#
// Developer Intent: Replace the Accent3 theme color with a specific green shade in every Excel workbook located in a given directory.
// Use Cases: Enforce corporate branding by updating the Accent3 color to the approved green across existing spreadsheets. | Prepare a collection of financial reports for a green‑themed presentation by applying the custom shade automatically. | Ensure design‑guideline compliance that mandates a particular green for Accent3 in all Excel files.
// AI Prompts: Generate C# code that uses Aspose.Cells to set Accent3 to #008000 for all .xlsx files in a folder, including error handling and logging. | Show how to read an RGB value from a JSON configuration file and apply it as the Accent3 theme color to multiple workbooks with Aspose.Cells. | Refactor the batch processing loop to use async file I/O while updating the Accent3 theme color in each workbook.

using System;
using System.IO;
using System.Drawing;
using Aspose.Cells;

namespace AsposeCellsThemeUpdater
{
    // Iterates through a specified folder, loads each .xlsx workbook with Aspose.Cells, changes the Accent3 theme color to a user‑defined green (RGB 0,128,0) via SetThemeColor, saves the file, and logs results while handling missing files and exceptions.
    class Program
    {
        static void Main()
        {
            // Folder containing the XLSX files – adjust as needed or keep empty to use the current directory.
            string folderPath = @"C:\Path\To\XlsxFolder";

            // Verify that the folder exists.
            if (!Directory.Exists(folderPath))
            {
                Console.WriteLine($"Folder not found: {folderPath}");
                Console.WriteLine("Please ensure the path is correct or update the folderPath variable.");
                return;
            }

            // Define the custom green shade to replace Accent3.
            Color customGreen = Color.FromArgb(0, 128, 0); // adjust RGB as needed

            try
            {
                // Iterate over each XLSX file in the folder.
                foreach (string filePath in Directory.GetFiles(folderPath, "*.xlsx"))
                {
                    // Ensure the file still exists before processing.
                    if (!File.Exists(filePath))
                    {
                        Console.WriteLine($"File not found (skipped): {filePath}");
                        continue;
                    }

                    try
                    {
                        // Load the workbook from the file (lifecycle rule: load).
                        Workbook workbook = new Workbook(filePath);

                        // Replace the Accent3 theme color with the custom green shade.
                        workbook.SetThemeColor(ThemeColorType.Accent3, customGreen);

                        // Save the workbook back to the same file (lifecycle rule: save).
                        workbook.Save(filePath, SaveFormat.Xlsx);

                        Console.WriteLine($"Processed: {Path.GetFileName(filePath)}");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error processing file '{filePath}': {ex.Message}");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }
    }
}
