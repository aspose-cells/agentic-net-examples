// Title: Batch apply a seasonal color theme to multiple quarterly report Excel workbooks using Aspose.Cells for .NET (C#)
// Common Searches: C# code to apply a custom background and font color to all worksheets in a folder of .xlsx files with Aspose.Cells | How to automate seasonal theming of quarterly Excel reports using Aspose.Cells batch processing | Aspose.Cells example for styling the used range of every worksheet in multiple workbooks | Programmatically set light green background for Excel workbooks in bulk with Aspose.Cells .NET | Processing multiple Excel files to add a green seasonal theme using Aspose.Cells in C#
// Tags: batch apply seasonal style Aspose.Cells .NET | style used range across multiple Excel workbooks C# | set worksheet background color programmatically Aspose.Cells | automate quarterly report theming Aspose.Cells | process multiple .xlsx files with Aspose.Cells batch job

using System;
using System.IO;
using Aspose.Cells;
using System.Drawing;

// Alias to avoid ambiguity with System.Range
using AsposeRange = Aspose.Cells.Range;

// The C# program scans a specified input directory for .xlsx quarterly reports, loads each workbook with Aspose.Cells, creates a light‑green background and dark‑green font style, applies the style to the used range of every worksheet, and saves the themed files to an output folder, handling missing files and runtime errors.
class SeasonalThemeBatchJob
{
    static void Main()
    {
        // Folder containing the quarterly report workbooks
        string inputFolder = @"C:\QuarterlyReports\Input";
        // Folder where the themed workbooks will be saved
        string outputFolder = @"C:\QuarterlyReports\Output";

        try
        {
            // Ensure the input directory exists
            if (!Directory.Exists(inputFolder))
            {
                Console.WriteLine($"Input folder does not exist: {inputFolder}");
                return;
            }

            // Ensure the output directory exists
            Directory.CreateDirectory(outputFolder);

            // Get all Excel files in the input folder (assumed to be .xlsx)
            string[] workbookFiles = Directory.GetFiles(inputFolder, "*.xlsx");

            foreach (string filePath in workbookFiles)
            {
                // Verify the file exists before loading
                if (!File.Exists(filePath))
                {
                    Console.WriteLine($"File not found, skipping: {filePath}");
                    continue;
                }

                try
                {
                    // Load the workbook
                    Workbook workbook = new Workbook(filePath);

                    // Seasonal colors
                    Color seasonalBackground = Color.FromArgb(220, 255, 230); // Light green
                    Color seasonalFontColor = Color.FromArgb(0, 100, 0);      // Dark green

                    // Apply style to each worksheet
                    foreach (Worksheet sheet in workbook.Worksheets)
                    {
                        // Determine used range
                        int maxRow = sheet.Cells.MaxDataRow;
                        int maxCol = sheet.Cells.MaxDataColumn;

                        // If the sheet is empty, skip styling
                        if (maxRow < 0 || maxCol < 0)
                            continue;

                        AsposeRange usedRange = sheet.Cells.CreateRange(0, 0, maxRow + 1, maxCol + 1);

                        // Create and configure style
                        Style seasonalStyle = workbook.CreateStyle();
                        seasonalStyle.ForegroundColor = seasonalBackground;
                        seasonalStyle.Pattern = BackgroundType.Solid;
                        seasonalStyle.Font.Color = seasonalFontColor;

                        // Apply style to the used range
                        StyleFlag flag = new StyleFlag { All = true };
                        usedRange.ApplyStyle(seasonalStyle, flag);
                    }

                    // Save the themed workbook to the output folder, preserving the original file name
                    string fileName = Path.GetFileName(filePath);
                    string outputPath = Path.Combine(outputFolder, fileName);
                    workbook.Save(outputPath);
                    Console.WriteLine($"Processed and saved: {outputPath}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing file '{filePath}': {ex.Message}");
                }
            }

            Console.WriteLine("Seasonal theme applied to all quarterly reports.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}
