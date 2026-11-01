// Title: Log per‑workbook conversion time while batch converting Excel files to HTML with Aspose.Cells for .NET
// AI Prompts: Wrap each Workbook.Save call with a timer and write the elapsed seconds to the console for every Excel file processed. | Update the batch conversion loop to capture and display the elapsed time even when an exception occurs, ensuring the timer stops before logging the error. | Create a reusable method that accepts an Excel file path, converts it to HTML using Aspose.Cells, returns the conversion time, and logs the result.
// Common Searches: how can I measure the time taken to convert each Excel workbook to HTML using Aspose.Cells in C# | C# batch export of .xlsx files to HTML with performance logging | Aspose.Cells SaveFormat.Html conversion duration per file example | record conversion time for multiple Excel files with a high‑resolution timer and Aspose.Cells | log errors and conversion time for each workbook during HTML export in .NET
// Tags: batch conversion timing Aspose.Cells | stopwatch logging Aspose.Cells HTML export | per‑file conversion duration .NET | error handling with conversion time Aspose.Cells | measure workbook save performance C#

using System;
using System.Diagnostics;
using System.IO;
using Aspose.Cells;

// The program scans a directory for .xlsx files, loads each workbook with Aspose.Cells, converts it to HTML, measures the conversion time using System.Diagnostics.Stopwatch, and writes the elapsed seconds to the console while handling missing files and conversion errors.
class BatchHtmlConverter
{
    static void Main(string[] args)
    {
        // Folder containing the source Excel workbooks
        string inputFolder = @"C:\InputExcels";

        // Folder where the generated HTML files will be saved
        string outputFolder = @"C:\OutputHtml";

        // Verify input folder exists
        if (!Directory.Exists(inputFolder))
        {
            Console.WriteLine($"Input folder not found: '{inputFolder}'.");
            return;
        }

        // Ensure the output directory exists
        Directory.CreateDirectory(outputFolder);

        // Retrieve all Excel files (adjust the pattern if needed)
        string[] excelFiles = Directory.GetFiles(inputFolder, "*.xlsx");

        foreach (string excelPath in excelFiles)
        {
            // Skip if the file somehow does not exist
            if (!File.Exists(excelPath))
            {
                Console.WriteLine($"File not found: '{excelPath}'. Skipping.");
                continue;
            }

            try
            {
                // Start measuring conversion time for the current workbook
                Stopwatch timer = Stopwatch.StartNew();

                // Load the workbook (lifecycle rule: load)
                Workbook workbook = new Workbook(excelPath);

                // Build the output HTML file path
                string htmlFileName = Path.GetFileNameWithoutExtension(excelPath) + ".html";
                string htmlPath = Path.Combine(outputFolder, htmlFileName);

                // Save the workbook as HTML (lifecycle rule: save)
                workbook.Save(htmlPath, SaveFormat.Html);

                // Stop the timer
                timer.Stop();

                // Log the duration of the conversion
                Console.WriteLine($"Workbook '{excelPath}' converted to HTML in {timer.Elapsed.TotalSeconds:F2} seconds.");
            }
            catch (Exception ex)
            {
                // Log any errors that occur during conversion
                Console.WriteLine($"Error converting '{excelPath}': {ex.Message}");
            }
        }
    }
}
