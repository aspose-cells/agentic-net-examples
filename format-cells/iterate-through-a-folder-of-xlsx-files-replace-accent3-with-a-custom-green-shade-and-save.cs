// Title: Batch replace Accent3 theme color with a custom green in XLSX files using Aspose.Cells C#
// Description: Iterates through a folder of .xlsx workbooks, reads the current theme, substitutes the Accent3 color with a user‑defined green RGB, applies a custom theme via Aspose.Cells, and saves each file. Includes folder validation and error handling.
// Keywords: Aspose.Cells C# replace theme color | Accent3 custom green | batch update Excel theme colors | CustomTheme method Aspose.Cells | process multiple XLSX files programmatically
// Common Searches: change Accent3 theme color in all Excel files with Aspose.Cells | C# code to replace Excel theme Accent3 with specific RGB | bulk update Excel workbook theme colors .NET | apply custom green theme to multiple XLSX files
// Developer Intent: Replace the Accent3 theme color with a custom green shade in every XLSX file within a specified directory and overwrite the originals.
// Use Cases: Enforce corporate green branding across all generated reports. | Prepare a batch of workbooks for brand‑compliance before distribution. | Automate a rebranding effort by updating the Accent3 color in existing Excel files.
// AI Prompts: Write C# code that uses Aspose.Cells to replace the Accent3 theme color with a given RGB value for all Excel files in a folder, adding error handling and optional backup creation. | Explain how Aspose.Cells' CustomTheme method can retrieve, modify, and reapply theme colors before saving a workbook. | Suggest how to make the folder path configurable via command‑line arguments and log processing results to a text file.

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;

// Iterates through a folder of .xlsx workbooks, reads the current theme, substitutes the Accent3 color with a user‑defined green RGB, applies a custom theme via Aspose.Cells, and saves each file. Includes folder validation and error handling.
class ReplaceAccent3WithCustomGreen
{
    static void Main()
    {
        // Folder containing the XLSX files – adjust as needed or pass via command line
        string folderPath = @"C:\Path\To\XlsxFolder";

        // Verify that the folder exists
        if (!Directory.Exists(folderPath))
        {
            Console.WriteLine($"Folder not found: {folderPath}");
            return;
        }

        // Define the custom green shade to replace Accent3
        Color customGreen = Color.FromArgb(0, 200, 0); // Adjust RGB as needed

        // Get all .xlsx files in the folder (non‑recursive)
        string[] files;
        try
        {
            files = Directory.GetFiles(folderPath, "*.xlsx", SearchOption.TopDirectoryOnly);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error retrieving files: {ex.Message}");
            return;
        }

        foreach (string filePath in files)
        {
            // Ensure the file still exists before processing
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"File not found (skipped): {filePath}");
                continue;
            }

            try
            {
                // Load the workbook from the file
                using (Workbook workbook = new Workbook(filePath))
                {
                    // Prepare an array for the 12 theme colors
                    Color[] themeColors = new Color[12];

                    // Fill the array with the current theme colors
                    themeColors[0] = workbook.GetThemeColor(ThemeColorType.Background1);
                    themeColors[1] = workbook.GetThemeColor(ThemeColorType.Text1);
                    themeColors[2] = workbook.GetThemeColor(ThemeColorType.Background2);
                    themeColors[3] = workbook.GetThemeColor(ThemeColorType.Text2);
                    themeColors[4] = workbook.GetThemeColor(ThemeColorType.Accent1);
                    themeColors[5] = workbook.GetThemeColor(ThemeColorType.Accent2);
                    themeColors[6] = workbook.GetThemeColor(ThemeColorType.Accent3); // will be replaced
                    themeColors[7] = workbook.GetThemeColor(ThemeColorType.Accent4);
                    themeColors[8] = workbook.GetThemeColor(ThemeColorType.Accent5);
                    themeColors[9] = workbook.GetThemeColor(ThemeColorType.Accent6);
                    themeColors[10] = workbook.GetThemeColor(ThemeColorType.Hyperlink);
                    themeColors[11] = workbook.GetThemeColor(ThemeColorType.FollowedHyperlink);

                    // Replace Accent3 (index 6) with the custom green shade
                    themeColors[6] = customGreen;

                    // Apply the modified theme to the workbook
                    workbook.CustomTheme("CustomThemeWithGreenAccent3", themeColors);

                    // Save the workbook, overwriting the original file
                    workbook.Save(filePath);
                }

                Console.WriteLine($"Processed: {Path.GetFileName(filePath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing file '{filePath}': {ex.Message}");
            }
        }

        Console.WriteLine("Accent3 color replacement completed for all files.");
    }
}
