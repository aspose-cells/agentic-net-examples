using System;
using System.IO;
using System.Drawing;
using Aspose.Cells;

namespace HolidayThemeBatch
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Folder containing the source workbooks
                string inputFolder = @"C:\Workbooks\Input";
                // Folder where the themed workbooks will be saved
                string outputFolder = @"C:\Workbooks\Output";

                // Verify input folder exists
                if (!Directory.Exists(inputFolder))
                {
                    Console.WriteLine($"Input folder not found: {inputFolder}");
                    return;
                }

                // Ensure the output directory exists
                Directory.CreateDirectory(outputFolder);

                // Define the 12 holiday theme colors
                Color[] holidayColors = new Color[]
                {
                    Color.FromArgb(255, 255, 255), // Background1 – White
                    Color.FromArgb(0, 0, 0),       // Text1 – Black
                    Color.FromArgb(255, 228, 196), // Background2 – Bisque
                    Color.FromArgb(139, 69, 19),   // Text2 – SaddleBrown
                    Color.FromArgb(255, 0, 0),     // Accent1 – Red
                    Color.FromArgb(0, 128, 0),     // Accent2 – Green
                    Color.FromArgb(255, 215, 0),   // Accent3 – Gold
                    Color.FromArgb(30, 144, 255),  // Accent4 – DodgerBlue
                    Color.FromArgb(255, 105, 180), // Accent5 – HotPink
                    Color.FromArgb(128, 0, 128),   // Accent6 – Purple
                    Color.FromArgb(0, 0, 255),     // Hyperlink – Blue
                    Color.FromArgb(128, 0, 0)      // Followed Hyperlink – Maroon
                };

                // Process each workbook in the input folder
                foreach (string inputPath in Directory.GetFiles(inputFolder, "*.xlsx"))
                {
                    // Verify the file exists before loading
                    if (!File.Exists(inputPath))
                    {
                        Console.WriteLine($"File not found: {inputPath}");
                        continue;
                    }

                    try
                    {
                        // Load the workbook
                        using (Workbook workbook = new Workbook(inputPath))
                        {
                            // Apply the custom holiday theme
                            workbook.CustomTheme("HolidayTheme", holidayColors);

                            // Determine the output file path
                            string fileName = Path.GetFileName(inputPath);
                            string outputPath = Path.Combine(outputFolder, fileName);

                            // Save the themed workbook
                            workbook.Save(outputPath);
                        }

                        Console.WriteLine($"Applied holiday theme to: {Path.GetFileName(inputPath)}");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error processing file '{inputPath}': {ex.Message}");
                    }
                }

                Console.WriteLine("Batch processing completed.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }
    }
}