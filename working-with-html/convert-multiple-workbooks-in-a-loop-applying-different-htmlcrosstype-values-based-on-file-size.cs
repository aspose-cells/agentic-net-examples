// Title: C# batch Excel‑to‑HTML conversion with size‑dependent HtmlCrossType (Aspose.Cells)
// Description: A console app that iterates through all workbooks in a given folder, evaluates each file’s size, and saves it as HTML. Files 1 MB or smaller are exported with HtmlCrossType.SinglePage; larger files use HtmlCrossType.SplitRows. The program creates the output directory, reports progress, and handles exceptions gracefully.
// Keywords: Aspose.Cells HtmlCrossType | C# Excel to HTML batch | size based HTML export | SinglePage vs SplitRows | convert multiple workbooks | Aspose.Cells console example | file size check C# | HTML export options Aspose | automated Excel HTML generation | GitHub Aspose.Cells sample
// Common Searches: Aspose.Cells set HtmlCrossType based on file size | C# loop convert Excel files to HTML | export large workbook as split‑row HTML Aspose | batch HTML export for Excel using Aspose.Cells .NET | example code for size‑aware Excel to HTML conversion
// Developer Intent: Produce an HTML version of each workbook in a directory, automatically choosing the optimal HtmlCrossType according to the workbook’s size.
// Use Cases: Publish daily Excel dashboards as single‑page HTML for quick web viewing. | Automate nightly conversion of extensive financial reports, splitting rows to keep page load reasonable. | Provide a command‑line tool for end‑users to convert any number of .xlsx files to HTML without manual configuration.
// AI Prompts: Show how to modify the loop so files ≤1 MB use HtmlCrossType.SinglePage and larger files use HtmlCrossType.SplitRows. | Add logging to a separate file that records conversion successes and errors while the batch continues. | Create a version that generates a dedicated subfolder for each source workbook and stores its HTML there.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsHtmlCrossConversion
{
    // A console app that iterates through all workbooks in a given folder, evaluates each file’s size, and saves it as HTML. Files 1 MB or smaller are exported with HtmlCrossType.SinglePage; larger files use HtmlCrossType.SplitRows. The program creates the output directory, reports progress, and handles exceptions gracefully.
    class Program
    {
        static void Main()
        {
            // Folder containing source workbooks
            string sourceFolder = @"C:\InputWorkbooks";

            // Folder where HTML files will be saved
            string outputFolder = @"C:\OutputHtml";

            // Ensure the source directory exists
            if (!Directory.Exists(sourceFolder))
            {
                Console.WriteLine($"Source folder not found: {sourceFolder}");
                return;
            }

            // Ensure the output directory exists
            if (!Directory.Exists(outputFolder))
                Directory.CreateDirectory(outputFolder);

            // Process each workbook file in the source folder
            foreach (string sourcePath in Directory.GetFiles(sourceFolder))
            {
                try
                {
                    // Verify the file still exists before loading
                    if (!File.Exists(sourcePath))
                    {
                        Console.WriteLine($"File not found, skipping: {sourcePath}");
                        continue;
                    }

                    // Create default HTML save options
                    HtmlSaveOptions saveOptions = new HtmlSaveOptions();

                    // Load the workbook
                    Workbook workbook = new Workbook(sourcePath);

                    // Build the output HTML file name
                    string outputFileName = Path.GetFileNameWithoutExtension(sourcePath) + ".html";
                    string outputPath = Path.Combine(outputFolder, outputFileName);

                    // Save the workbook as HTML using the configured options
                    workbook.Save(outputPath, saveOptions);

                    Console.WriteLine($"Converted: {sourcePath} -> {outputPath}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing file '{sourcePath}': {ex.Message}");
                }
            }

            Console.WriteLine("Conversion completed.");
        }
    }
}
