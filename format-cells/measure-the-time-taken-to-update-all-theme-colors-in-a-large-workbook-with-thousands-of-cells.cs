// Title: Measure time taken to update all theme colors in a large Excel workbook with Aspose.Cells for .NET
// Common Searches: how to benchmark updating theme colors in a large .xlsx file using Aspose.Cells C# | C# code to time theme color inversion across an entire workbook with Aspose.Cells | measure execution time of Excel theme color changes with Stopwatch and Aspose.Cells | performance testing for updating all theme colors in a big workbook via Aspose.Cells | using Aspose.Cells .NET to profile theme color updates in a large Excel spreadsheet
// Tags: measure theme color update time Aspose.Cells | update workbook theme colors C# | performance timing Excel theme modification .NET | large workbook theme color inversion Aspose.Cells | stopwatch benchmark Aspose.Cells operations

using System;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using Aspose.Cells;

// The example loads a large Excel workbook, starts a Stopwatch, accesses the workbook's theme color scheme via dynamic, inverts each of the theme colors, stops the timer, outputs the elapsed milliseconds, and saves the modified workbook.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "LargeWorkbook.xlsx";
            string outputPath = "LargeWorkbook_Updated.xlsx";

            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Start measuring time
            Stopwatch stopwatch = Stopwatch.StartNew();

            try
            {
                // Access theme and color scheme using dynamic to avoid compile‑time type dependencies
                dynamic theme = workbook.Theme;
                dynamic colorScheme = theme.ColorScheme;

                // Update each theme color (typically 12 colors)
                int count = (int)colorScheme.Count;
                for (int i = 0; i < count; i++)
                {
                    Color oldColor = (Color)colorScheme[i];
                    Color newColor = Color.FromArgb(255 - oldColor.R, 255 - oldColor.G, 255 - oldColor.B);
                    colorScheme[i] = newColor;
                }
            }
            catch (Exception ex)
            {
                // If the workbook does not support themes, log and continue
                Console.WriteLine($"Theme update skipped: {ex.Message}");
            }

            // Stop measuring time
            stopwatch.Stop();

            // Output the elapsed time
            Console.WriteLine($"Time taken to update theme colors (if applicable): {stopwatch.ElapsedMilliseconds} ms");

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
