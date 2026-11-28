// Title: Log per‑workbook conversion time while batch converting Excel files to HTML with Aspose.Cells for .NET
// Common Searches: C# how to record conversion time for each Excel workbook when exporting to HTML using Aspose.Cells | Aspose.Cells batch HTML export log duration per file example | measure Excel to HTML conversion performance in a .NET console application | save conversion timestamps to a log file during Aspose.Cells batch processing | track processing time of multiple workbooks with Stopwatch and Aspose.Cells
// Tags: Aspose.Cells batch HTML conversion timing | C# Stopwatch Excel to HTML performance logging | per‑file conversion duration logging Aspose.Cells | log workbook processing time .NET | measure Excel workbook HTML export time

using System;
using System.IO;
using System.Diagnostics;
using Aspose.Cells;

namespace BatchHtmlConversion
{
    // The console program iterates over Excel workbooks in a folder, converts each to HTML using Aspose.Cells' HtmlSaveOptions, measures the conversion duration with a Stopwatch, and writes the elapsed seconds to both the console and a log file, handling errors and summarizing the batch results.
    class Program
    {
        static int Main(string[] args)
        {
            // Resolve arguments or fall back to defaults
            string inputFolder = args.Length > 0 ? args[0] : "InputWorkbooks";
            string outputFolder = args.Length > 1 ? args[1] : "OutputHtml";
            string logFilePath = args.Length > 2 ? args[2] : "conversion_log.txt";

            // Ensure directories exist
            if (!Directory.Exists(inputFolder))
            {
                Console.Error.WriteLine($"Input folder does not exist: {inputFolder}");
                return 1;
            }

            Directory.CreateDirectory(outputFolder);
            Directory.CreateDirectory(Path.GetDirectoryName(Path.GetFullPath(logFilePath)) ?? ".");

            // Prepare the log file (overwrite any existing file)
            File.WriteAllText(logFilePath, $"Batch HTML conversion started at {DateTime.Now:O}{Environment.NewLine}");

            // Supported workbook extensions
            string[] extensions = new[] { ".xlsx", ".xls", ".xlsm", ".xlsb" };
            string[] workbookFiles = Directory.GetFiles(inputFolder, "*.*", SearchOption.TopDirectoryOnly);
            int processedCount = 0;

            foreach (string workbookPath in workbookFiles)
            {
                if (Array.IndexOf(extensions, Path.GetExtension(workbookPath).ToLowerInvariant()) < 0)
                {
                    // Skip non‑workbook files
                    continue;
                }

                string fileName = Path.GetFileName(workbookPath);
                string htmlFileName = Path.ChangeExtension(fileName, ".html");
                string htmlOutputPath = Path.Combine(outputFolder, htmlFileName);

                Stopwatch sw = new Stopwatch();
                try
                {
                    sw.Start();

                    // Load the workbook
                    Workbook workbook = new Workbook(workbookPath);

                    // Configure HTML save options (customize as needed)
                    HtmlSaveOptions htmlOptions = new HtmlSaveOptions
                    {
                        ExportGridLines = true,
                        ExportActiveWorksheetOnly = false,
                        // You can enable PresentationPreference for nicer output if using GridWeb options,
                        // but HtmlSaveOptions does not expose it directly.
                    };

                    // Save as HTML
                    workbook.Save(htmlOutputPath, htmlOptions);

                    sw.Stop();

                    string logEntry = $"{fileName} => {htmlFileName} | Duration: {sw.Elapsed.TotalSeconds:F3} sec";
                    Console.WriteLine(logEntry);
                    File.AppendAllText(logFilePath, logEntry + Environment.NewLine);
                    processedCount++;
                }
                catch (Exception ex)
                {
                    sw.Stop();
                    string errorEntry = $"ERROR processing {fileName}: {ex.Message}";
                    Console.Error.WriteLine(errorEntry);
                    File.AppendAllText(logFilePath, errorEntry + Environment.NewLine);
                }
            }

            string summary = $"Batch completed. Processed {processedCount} workbook(s). Log saved to {Path.GetFullPath(logFilePath)}";
            Console.WriteLine(summary);
            File.AppendAllText(logFilePath, summary + Environment.NewLine);

            return 0;
        }
    }
}
