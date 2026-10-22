// Title: Extract Excel workbooks from a ZIP archive and convert each to HTML with Aspose.Cells for .NET
// AI Prompts: Write a C# console application that opens a zip file, iterates over every .xls or .xlsx entry, loads each workbook with Aspose.Cells, and saves it as an HTML file using SaveFormat.Html. | Extend the program to accept command‑line parameters for the zip path and output directory, and generate a log file that records the conversion result for each workbook.
// Common Searches: how to batch convert .xlsx files inside a zip to html using Aspose.Cells C# | C# extract Excel workbooks from a zip archive and export to html | Aspose.Cells convert multiple Excel files to html in a loop .NET
// Tags: Aspose.Cells unzip Excel to HTML | C# batch convert Excel workbooks to HTML | SaveFormat.Html workbook export | process zip entries with Aspose.Cells | convert .xlsx stream to HTML C#

using System;
using System.IO;
using System.IO.Compression;
using Aspose.Cells;

// The program opens a ZIP archive, extracts each .xls/.xlsx entry, loads it into an Aspose.Cells Workbook, and saves the workbook as an HTML file in a specified output folder.
class Program
{
    static void Main(string[] args)
    {
        // Path to the zip archive containing Excel files
        string zipPath = "input.zip";

        // Verify that the zip file exists
        if (!File.Exists(zipPath))
        {
            Console.WriteLine($"Error: Zip file '{zipPath}' not found.");
            return;
        }

        // Directory where the generated HTML files will be saved
        string outputDirectory = "output_html";

        // Ensure the output directory exists
        Directory.CreateDirectory(outputDirectory);

        try
        {
            // Open the zip archive for reading
            using (ZipArchive archive = ZipFile.OpenRead(zipPath))
            {
                // Iterate through each entry in the archive
                foreach (ZipArchiveEntry entry in archive.Entries)
                {
                    // Process only files with .xls or .xlsx extensions
                    if (entry.FullName.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase) ||
                        entry.FullName.EndsWith(".xls", StringComparison.OrdinalIgnoreCase))
                    {
                        // Open a stream for the current Excel entry
                        using (Stream entryStream = entry.Open())
                        {
                            // Load the workbook from the stream
                            Workbook workbook = new Workbook(entryStream);

                            // Determine the output HTML file name
                            string htmlFileName = Path.GetFileNameWithoutExtension(entry.Name) + ".html";
                            string htmlPath = Path.Combine(outputDirectory, htmlFileName);

                            // Save the workbook as an HTML file
                            workbook.Save(htmlPath, SaveFormat.Html);
                        }
                    }
                }
            }

            Console.WriteLine("All Excel files have been converted to HTML.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
