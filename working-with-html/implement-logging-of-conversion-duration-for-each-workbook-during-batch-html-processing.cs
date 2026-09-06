// Title: Log the time taken to convert each Excel workbook to HTML in a batch process using Aspose.Cells for .NET
// AI Prompts: Wrap the workbook.Save call with a Stopwatch and write the elapsed seconds to the console for every file in the loop. | Create a helper method that accepts an Excel file path, converts it to HTML with Aspose.Cells, returns the conversion duration, and logs the result. | Add try‑catch around each conversion so that the elapsed time is logged even when a workbook fails to save.
// Common Searches: how to measure per‑file conversion time when using Aspose.Cells to save Excel as HTML in C# | C# batch convert .xlsx to .html and log duration for each workbook | Aspose.Cells performance logging for multiple workbook HTML exports | record elapsed time for each workbook.Save operation in a .NET console app
// Tags: batch workbook to HTML conversion timing with Aspose.Cells | C# Stopwatch logging for Aspose.Cells save operation | measure per‑file conversion duration .NET | Aspose.Cells HTML export performance tracking

using System;
using System.IO;
using System.Diagnostics;
using Aspose.Cells;

// The program iterates through all .xlsx files in a folder, loads each workbook with Aspose.Cells, converts it to HTML, measures the conversion time using Stopwatch, and writes the elapsed seconds to the console for every workbook.
class Program
{
    static void Main(string[] args)
    {
        // Define input and output directories
        string inputDir = @"C:\InputWorkbooks";
        string outputDir = @"C:\OutputHtml";

        // Ensure the output directory exists
        Directory.CreateDirectory(outputDir);

        // Retrieve all Excel files from the input directory
        string[] excelFiles = Directory.GetFiles(inputDir, "*.xlsx");

        foreach (string excelPath in excelFiles)
        {
            // Load the workbook (create/load rule)
            Workbook workbook = new Workbook(excelPath);

            // Start timing the conversion
            Stopwatch timer = Stopwatch.StartNew();

            // Define the output HTML file path
            string htmlPath = Path.Combine(outputDir, Path.GetFileNameWithoutExtension(excelPath) + ".html");

            // Save the workbook as HTML (save rule)
            workbook.Save(htmlPath, SaveFormat.Html);

            // Stop timing
            timer.Stop();

            // Log the conversion duration
            Console.WriteLine($"Converted '{Path.GetFileName(excelPath)}' to HTML in {timer.Elapsed.TotalSeconds:F2} seconds.");
        }
    }
}
