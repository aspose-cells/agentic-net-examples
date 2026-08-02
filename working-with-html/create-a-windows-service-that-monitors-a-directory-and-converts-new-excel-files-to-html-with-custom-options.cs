// Title: C# Windows Service to Watch a Folder and Convert Excel Workbooks to Single‑File HTML with Aspose.Cells
// Description: A self‑contained Windows service that uses FileSystemWatcher to detect new Excel files (.xlsx, .xls, .xlsm, .xlsb) in a specified directory, loads each workbook with Aspose.Cells, applies HtmlSaveOptions (HTML5, gridlines, export all worksheets, single‑file output, custom page title), and writes the resulting HTML to a target folder. The service can be started and stopped programmatically and includes retry logic for file‑access readiness.
// Keywords: Aspose.Cells Excel to HTML conversion | C# FileSystemWatcher folder monitor | Windows service convert Excel to HTML | HtmlSaveOptions single file Aspose | Export gridlines HTML5 Aspose.Cells | automatic Excel HTML generation | C# background service Excel conversion
// Common Searches: how to create a Windows service that converts Excel to HTML using Aspose.Cells | C# folder watcher that converts .xlsx files to single HTML file | Aspose.Cells HtmlSaveOptions example for HTML5 and gridlines | automated Excel to HTML conversion service .NET | file system watcher retry logic for Excel file processing
// Developer Intent: Build a background Windows service that continuously watches a folder and transforms any newly added Excel workbook into a single HTML file with custom formatting options.
// Use Cases: Generate web‑ready reports instantly when users drop Excel files into a shared drop folder. | Publish financial or operational spreadsheets as intranet HTML pages without manual steps. | Provide on‑the‑fly HTML previews for a document‑management system by converting uploaded Excel files.
// AI Prompts: Create a C# Windows Service that monitors a directory and uses Aspose.Cells to convert new Excel files to a single HTML5 file with gridlines and a custom page title. | Write unit tests for the ConvertToHtml method to verify HtmlSaveOptions settings and confirm the HTML output is created in the correct location. | Explain how to enhance the FileSystemWatcher to ignore temporary Office files (e.g., ~$.xlsx) and safely handle rename and change events.

using System;
using System.IO;
using System.Threading;
using Aspose.Cells;

namespace ExcelToHtmlService
{
    // A self‑contained Windows service that uses FileSystemWatcher to detect new Excel files (.xlsx, .xls, .xlsm, .xlsb) in a specified directory, loads each workbook with Aspose.Cells, applies HtmlSaveOptions (HTML5, gridlines, export all worksheets, single‑file output, custom page title), and writes the resulting HTML to a target folder. The service can be started and stopped programmatically and includes retry logic for file‑access readiness.
    public class ConverterService
    {
        private FileSystemWatcher _watcher;
        private readonly string _inputFolder = @"C:\WatchFolder";
        private readonly string _outputFolder = @"C:\HtmlOutput";

        // Starts the watcher and ensures output folder exists
        public void OnStart(string[] args)
        {
            try
            {
                Directory.CreateDirectory(_outputFolder);

                _watcher = new FileSystemWatcher(_inputFolder)
                {
                    Filter = "*.*",
                    EnableRaisingEvents = true,
                    IncludeSubdirectories = false
                };
                _watcher.Created += OnCreated;
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Failed to start service: {ex.Message}");
            }
        }

        // Stops the watcher and releases resources
        public void OnStop()
        {
            try
            {
                if (_watcher != null)
                {
                    _watcher.EnableRaisingEvents = false;
                    _watcher.Created -= OnCreated;
                    _watcher.Dispose();
                    _watcher = null;
                }
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Failed to stop service: {ex.Message}");
            }
        }

        // Handles new files; retries until the file is ready
        private void OnCreated(object sender, FileSystemEventArgs e)
        {
            for (int i = 0; i < 5; i++)
            {
                try
                {
                    using (FileStream stream = File.Open(e.FullPath, FileMode.Open, FileAccess.Read, FileShare.None))
                    {
                        // File opened exclusively – ready for processing
                        break;
                    }
                }
                catch
                {
                    Thread.Sleep(500);
                }
            }

            string ext = Path.GetExtension(e.FullPath).ToLowerInvariant();
            if (ext == ".xlsx" || ext == ".xls" || ext == ".xlsm" || ext == ".xlsb")
            {
                try
                {
                    ConvertToHtml(e.FullPath);
                }
                catch (Exception ex)
                {
                    Console.Error.WriteLine($"Error converting '{e.FullPath}': {ex.Message}");
                }
            }
        }

        // Converts an Excel workbook to a single HTML file
        private void ConvertToHtml(string sourcePath)
        {
            if (!File.Exists(sourcePath))
            {
                Console.Error.WriteLine($"Source file not found: {sourcePath}");
                return;
            }

            try
            {
                // Load workbook
                Workbook workbook = new Workbook(sourcePath);

                // Set HTML save options
                HtmlSaveOptions saveOptions = new HtmlSaveOptions
                {
                    ExportActiveWorksheetOnly = false,
                    ExportGridLines = true,
                    HtmlVersion = HtmlVersion.Html5,
                    PageTitle = Path.GetFileNameWithoutExtension(sourcePath),
                    SaveAsSingleFile = true
                };

                // Determine destination path
                string destFileName = Path.GetFileNameWithoutExtension(sourcePath) + ".html";
                string destPath = Path.Combine(_outputFolder, destFileName);

                // Save as HTML
                workbook.Save(destPath, saveOptions);
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Failed to convert '{sourcePath}': {ex.Message}");
            }
        }

        // Entry point – runs as a console application
        public static void Main(string[] args)
        {
            ConverterService service = new ConverterService();
            service.OnStart(args);
            Console.WriteLine("Service started. Press Enter to stop...");
            Console.ReadLine();
            service.OnStop();
        }
    }
}
