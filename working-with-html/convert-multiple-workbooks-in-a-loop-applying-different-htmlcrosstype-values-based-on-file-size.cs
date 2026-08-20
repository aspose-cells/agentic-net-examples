// Title: Batch Convert Excel Workbooks to HTML with Size‑Based HtmlCrossType (Aspose.Cells .NET)
// Description: Iterates through a folder of Excel files, skips unsupported formats, checks each file's size, applies HtmlCrossType.Cross for files larger than 5 MB and HtmlCrossType.Default for smaller ones, and saves the workbooks as HTML using Aspose.Cells HtmlSaveOptions. Logs each conversion and creates the output directory if needed.
// Keywords: Aspose.Cells | C# | .NET | HtmlCrossType | HtmlSaveOptions | batch Excel to HTML | size based conversion | Excel to HTML loop | cross type HTML export | large workbook HTML rendering
// Common Searches: Aspose.Cells set HtmlCrossType by file size | batch convert Excel to HTML C# Aspose | HtmlCrossStringType conditional example | process multiple workbooks and save as HTML | how to use HtmlCrossType.Cross for large Excel files
// Developer Intent: Convert a collection of Excel workbooks to HTML, automatically choosing HtmlCrossType.Cross for files over 5 MB and HtmlCrossType.Default for smaller files.
// Use Cases: Automated nightly job that publishes financial spreadsheets as web‑ready HTML, optimizing large files with cross‑type rendering. | Migration script that prepares mixed‑size Excel reports for a web portal, ensuring appropriate HTML output size and performance. | Bulk generation of documentation from a repository of Excel templates, applying size‑aware HtmlCrossType to control HTML payload.
// AI Prompts: Create a C# utility that accepts input and output folder paths and converts all Excel files to HTML, using HtmlCrossType.Cross for files >5 MB and HtmlCrossType.Default otherwise with Aspose.Cells. | Refactor the sample to extract the size‑based HtmlCrossType logic into a reusable method and add comprehensive error handling for missing files and permission issues. | Write unit tests that mock FileInfo.Length to verify the correct HtmlCrossType is applied during batch HTML conversion using Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsHtmlCrossDemo
{
    // Iterates through a folder of Excel files, skips unsupported formats, checks each file's size, applies HtmlCrossType.Cross for files larger than 5 MB and HtmlCrossType.Default for smaller ones, and saves the workbooks as HTML using Aspose.Cells HtmlSaveOptions. Logs each conversion and creates the output directory if needed.
    class Program
    {
        static void Main()
        {
            // Directory containing source Excel files
            string sourceDir = @"C:\InputWorkbooks";
            // Directory where HTML files will be saved
            string outputDir = @"C:\OutputHtml";

            // Ensure output directory exists
            Directory.CreateDirectory(outputDir);

            // Process each Excel file in the source directory
            foreach (string filePath in Directory.GetFiles(sourceDir, "*.*", SearchOption.TopDirectoryOnly))
            {
                // Load only supported Excel formats
                string ext = Path.GetExtension(filePath).ToLowerInvariant();
                if (ext != ".xlsx" && ext != ".xls" && ext != ".xlsm" && ext != ".xlsb")
                    continue;

                // Determine file size
                long fileSize = new FileInfo(filePath).Length;

                // Choose HtmlCrossType based on size
                HtmlCrossType crossType = fileSize > 5 * 1024 * 1024   // >5 MB
                    ? HtmlCrossType.Cross
                    : HtmlCrossType.Default;

                // Load workbook
                using (Workbook workbook = new Workbook(filePath))
                {
                    // Configure HTML save options
                    HtmlSaveOptions saveOptions = new HtmlSaveOptions();
                    saveOptions.HtmlCrossStringType = crossType;

                    // Build output HTML file name
                    string outputFileName = Path.Combine(outputDir,
                        Path.GetFileNameWithoutExtension(filePath) + ".html");

                    // Save workbook as HTML with the specified options
                    workbook.Save(outputFileName, saveOptions);
                }

                Console.WriteLine($"Converted '{Path.GetFileName(filePath)}' to HTML using HtmlCrossType.{crossType}");
            }

            Console.WriteLine("All files have been processed.");
        }
    }
}
