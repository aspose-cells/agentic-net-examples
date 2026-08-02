// Title: Batch Synchronize Excel Theme Colors Across Workbooks with Aspose.Cells for .NET (C#)
// Description: A C# console utility that scans a folder of .xlsx files, uses the first workbook as a reference, compares the 12 theme colors of each file, and automatically copies the reference theme to any workbook with mismatched colors. Ideal for enforcing corporate branding or updating legacy reports in bulk.
// Keywords: Aspose.Cells theme synchronization | C# batch update Excel theme colors | compare Excel theme colors .NET | copy theme between workbooks Aspose | detect mismatched theme colors C# | Excel branding automation | bulk Excel theme alignment | Aspose.Cells GetThemeColor | CopyTheme method example | Excel workbook theme consistency
// Common Searches: how to copy Excel theme from one workbook to another using Aspose.Cells | C# code to compare and sync theme colors across multiple .xlsx files | batch update Excel theme colors with Aspose.Cells .NET | detect mismatched theme colors in Excel workbooks programmatically | Aspose.Cells example for theme synchronization
// Developer Intent: Automatically align the theme colors of all Excel files in a directory with a designated reference workbook.
// Use Cases: Enforce corporate brand colors across all generated reports. | Refresh legacy spreadsheets after a visual identity change. | Validate that exported Excel files meet a predefined theme before distribution. | Integrate theme compliance into a CI/CD pipeline for data‑export tools.
// AI Prompts: Write a C# console program using Aspose.Cells that reads a folder of .xlsx files, extracts the 12 theme colors from the first file, compares them with each subsequent file, and copies the reference theme to any workbook where the colors differ. | Generate a method called SyncWorkbookThemes(string folderPath) that returns a list of files updated after copying the reference theme using Aspose.Cells. | Provide a step‑by‑step explanation of how to use Workbook.GetThemeColor and Workbook.CopyTheme to ensure theme consistency across multiple Excel workbooks.

using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using Aspose.Cells;

namespace ThemeSynchronizationDemo
{
    // A C# console utility that scans a folder of .xlsx files, uses the first workbook as a reference, compares the 12 theme colors of each file, and automatically copies the reference theme to any workbook with mismatched colors. Ideal for enforcing corporate branding or updating legacy reports in bulk.
    class Program
    {
        static void Main()
        {
            try
            {
                // Folder that contains the workbooks to be checked.
                string folderPath = @"C:\ExcelFiles";

                // Verify that the folder exists.
                if (!Directory.Exists(folderPath))
                {
                    Console.WriteLine($"Folder not found: {folderPath}");
                    return;
                }

                // Get all .xlsx files in the folder.
                string[] excelFiles = Directory.GetFiles(folderPath, "*.xlsx", SearchOption.TopDirectoryOnly);
                if (excelFiles.Length == 0)
                {
                    Console.WriteLine("No Excel files found in the specified folder.");
                    return;
                }

                // Load the first workbook – it will serve as the reference theme source.
                string referencePath = excelFiles[0];
                if (!File.Exists(referencePath))
                {
                    Console.WriteLine($"Reference file not found: {referencePath}");
                    return;
                }

                using (Workbook referenceWorkbook = new Workbook(referencePath))
                {
                    // Store the reference theme colors for quick comparison.
                    Dictionary<ThemeColorType, Color> referenceColors = new Dictionary<ThemeColorType, Color>();
                    foreach (ThemeColorType type in Enum.GetValues(typeof(ThemeColorType)))
                    {
                        // ThemeColorType has values beyond the 12 theme colors; we limit to 0‑11.
                        if ((int)type > 11) continue;
                        referenceColors[type] = referenceWorkbook.GetThemeColor(type);
                    }

                    // Process each workbook other than the reference.
                    for (int i = 1; i < excelFiles.Length; i++)
                    {
                        string filePath = excelFiles[i];
                        if (!File.Exists(filePath))
                        {
                            Console.WriteLine($"File not found, skipping: {filePath}");
                            continue;
                        }

                        try
                        {
                            using (Workbook targetWorkbook = new Workbook(filePath))
                            {
                                bool mismatched = false;

                                // Compare each of the 12 theme colors.
                                foreach (ThemeColorType type in referenceColors.Keys)
                                {
                                    Color targetColor = targetWorkbook.GetThemeColor(type);
                                    if (targetColor.ToArgb() != referenceColors[type].ToArgb())
                                    {
                                        mismatched = true;
                                        break;
                                    }
                                }

                                if (mismatched)
                                {
                                    // Synchronize the theme by copying from the reference workbook.
                                    targetWorkbook.CopyTheme(referenceWorkbook);

                                    // Overwrite the original file with the synchronized theme.
                                    targetWorkbook.Save(filePath, SaveFormat.Xlsx);
                                    Console.WriteLine($"Synchronized theme for: {Path.GetFileName(filePath)}");
                                }
                                else
                                {
                                    Console.WriteLine($"Theme already matches for: {Path.GetFileName(filePath)}");
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"Error processing file '{Path.GetFileName(filePath)}': {ex.Message}");
                        }
                    }
                }

                Console.WriteLine("Theme synchronization completed.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }
    }
}
