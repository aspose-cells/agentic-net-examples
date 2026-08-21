// Title: Windows Service that Monitors a Folder and Converts Excel to HTML with Aspose.Cells
// Description: A C# Windows service that watches a designated directory, detects newly added Excel workbooks (xlsx, xls, xlsm, xlsb), and instantly converts them to a single HTML5 file using Aspose.Cells. The conversion applies custom HtmlSaveOptions such as grid‑line display, base‑64 image embedding, and export of all worksheets, while providing start/stop control and basic error handling.
// Keywords: Aspose.Cells | C# | Excel to HTML conversion | FileSystemWatcher | Windows Service | HtmlSaveOptions | HTML5 output | embed images base64 | export grid lines | background file conversion
// Common Searches: C# Windows service to convert Excel files to HTML | Aspose.Cells HtmlSaveOptions single file with embedded images | FileSystemWatcher trigger Excel to HTML conversion | automate Excel to HTML conversion in .NET | monitor folder and generate HTML from Excel
// Developer Intent: Create a background Windows service that automatically transforms incoming Excel files into self‑contained HTML pages using Aspose.Cells with specific formatting options.
// Use Cases: Generate web‑ready reports as soon as users drop Excel files into a shared folder. | Archive Excel workbooks as single HTML5 pages for intranet publishing or email distribution. | Run a low‑maintenance service that converts batch Excel uploads into self‑contained HTML for downstream systems.
// AI Prompts: Provide a complete Windows Service project template (including installer) that integrates the ExcelToHtmlService class and logs conversions with Serilog. | Write NUnit tests for ConvertExcelToHtml covering missing file, load failure, and save failure scenarios. | Suggest enhancements such as configurable watch paths via appsettings.json, retry logic for file locks, and optional PDF fallback conversion.

using System;
using System.IO;
using System.Threading;
using Aspose.Cells;

namespace ExcelToHtmlWindowsService
{
    // Simple class that watches a folder and converts new Excel files to HTML
    // A C# Windows service that watches a designated directory, detects newly added Excel workbooks (xlsx, xls, xlsm, xlsb), and instantly converts them to a single HTML5 file using Aspose.Cells. The conversion applies custom HtmlSaveOptions such as grid‑line display, base‑64 image embedding, and export of all worksheets, while providing start/stop control and basic error handling.
    public class ExcelToHtmlService
    {
        private FileSystemWatcher _watcher;
        private readonly string _sourceFolder = @"C:\WatchedFolder";
        private readonly string _outputFolder = @"C:\HtmlOutput";

        // Starts the watcher
        public void Start(string[] args)
        {
            // Ensure source and output directories exist
            if (!Directory.Exists(_sourceFolder))
                Directory.CreateDirectory(_sourceFolder);

            if (!Directory.Exists(_outputFolder))
                Directory.CreateDirectory(_outputFolder);

            try
            {
                // Set up the file system watcher
                _watcher = new FileSystemWatcher(_sourceFolder)
                {
                    Filter = "*.*",
                    NotifyFilter = NotifyFilters.FileName | NotifyFilters.CreationTime,
                    EnableRaisingEvents = true
                };

                // Subscribe to created event
                _watcher.Created += OnCreated;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to start watcher: {ex.Message}");
                throw;
            }
        }

        // Stops the watcher and releases resources
        public void Stop()
        {
            if (_watcher != null)
            {
                _watcher.EnableRaisingEvents = false;
                _watcher.Created -= OnCreated;
                _watcher.Dispose();
                _watcher = null;
            }
        }

        // Event handler for new files
        private void OnCreated(object sender, FileSystemEventArgs e)
        {
            // Small delay to ensure the file is fully written
            Thread.Sleep(500);

            // Process only Excel files
            string ext = Path.GetExtension(e.FullPath).ToLowerInvariant();
            if (ext == ".xlsx" || ext == ".xls" || ext == ".xlsm" || ext == ".xlsb")
            {
                try
                {
                    ConvertExcelToHtml(e.FullPath);
                }
                catch (Exception ex)
                {
                    // Log or handle exception as needed
                    Console.WriteLine($"Error converting '{e.Name}': {ex.Message}");
                }
            }
        }

        // Core conversion logic using Aspose.Cells APIs
        private void ConvertExcelToHtml(string excelPath)
        {
            // Verify the source file exists before loading
            if (!File.Exists(excelPath))
                throw new FileNotFoundException("Excel file not found.", excelPath);

            Workbook workbook;
            try
            {
                // Load the workbook
                workbook = new Workbook(excelPath);
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"Failed to load workbook '{excelPath}'.", ex);
            }

            // Configure HTML save options
            HtmlSaveOptions saveOptions = new HtmlSaveOptions
            {
                ExportActiveWorksheetOnly = false, // Export all worksheets
                ExportGridLines = true,            // Show grid lines
                HtmlVersion = HtmlVersion.Html5,   // Use HTML5 standard
                SaveAsSingleFile = true,           // Single HTML file
                ExportImagesAsBase64 = true        // Embed images
            };

            // Determine output file name
            string fileNameWithoutExt = Path.GetFileNameWithoutExtension(excelPath);
            string htmlPath = Path.Combine(_outputFolder, fileNameWithoutExt + ".html");

            try
            {
                // Save the workbook as HTML
                workbook.Save(htmlPath, saveOptions);
                Console.WriteLine($"Converted '{excelPath}' to '{htmlPath}'.");
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"Failed to save HTML file '{htmlPath}'.", ex);
            }
        }

        // Entry point for debugging as console app
        public static void Main(string[] args)
        {
            ExcelToHtmlService service = new ExcelToHtmlService();
            try
            {
                service.Start(args);
                Console.WriteLine("Service started. Press Enter to stop...");
                Console.ReadLine();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Service failed to start: {ex.Message}");
            }
            finally
            {
                service.Stop();
                Console.WriteLine("Service stopped.");
            }
        }
    }
}
