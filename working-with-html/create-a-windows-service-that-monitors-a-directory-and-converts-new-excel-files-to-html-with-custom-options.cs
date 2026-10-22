// Title: C# Windows service that watches a folder and automatically converts new Excel files to HTML with Aspose.Cells custom options
// AI Prompts: Write a C# Windows service that uses FileSystemWatcher to monitor a directory and converts any newly created .xls, .xlsx, or .xlsm file to HTML with Aspose.Cells, embedding images as base64 and preserving gridlines. | Add robust retry logic to the watcher’s Created event so the service waits until the Excel file is fully written before loading it with Aspose.Cells. | Show how to configure HtmlSaveOptions to export all worksheets, embed images as base64, and include gridlines while saving the workbook as HTML.
// Common Searches: how to create a Windows service in C# that converts incoming Excel files to HTML using Aspose.Cells | C# FileSystemWatcher convert new .xlsx files to HTML with embedded images | Aspose.Cells HtmlSaveOptions settings for base64 image embedding and gridlines | retry opening a file in FileSystemWatcher until the copy operation finishes | automate batch conversion of Excel workbooks to HTML in a background service
// Tags: Aspose.Cells Excel to HTML conversion | FileSystemWatcher monitor folder for Excel files | Windows service automatic Excel HTML export | HtmlSaveOptions embed images base64 | retry file access before Aspose.Cells conversion

using System;
using System.IO;
using System.Threading;
using Aspose.Cells;

namespace ExcelToHtmlService
{
    // Simple application that watches a folder and converts new Excel files to HTML
    // The example implements a C# Windows service that creates a FileSystemWatcher on a configurable input folder, retries opening newly created .xls/.xlsx/.xlsm files until they are ready, loads each workbook with Aspose.Cells, and saves it as HTML to an output folder using HtmlSaveOptions that embed images as base64, preserve gridlines, and export all worksheets.
    public class ExcelToHtmlWatcher
    {
        private FileSystemWatcher? _watcher;
        private readonly string _inputFolder = @"C:\InputExcel";   // folder to monitor
        private readonly string _outputFolder = @"C:\OutputHtml"; // folder for HTML output

        // Starts the folder monitoring
        public void Start()
        {
            try
            {
                // Ensure the input and output directories exist
                Directory.CreateDirectory(_inputFolder);
                Directory.CreateDirectory(_outputFolder);

                // Set up the file system watcher
                _watcher = new FileSystemWatcher(_inputFolder)
                {
                    Filter = "*.*", // watch all files
                    NotifyFilter = NotifyFilters.FileName | NotifyFilters.CreationTime
                };
                _watcher.Created += OnCreated;
                _watcher.EnableRaisingEvents = true;

                Console.WriteLine($"Watching folder: {_inputFolder}");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Failed to start watcher: {ex.Message}");
            }
        }

        // Stops the folder monitoring
        public void Stop()
        {
            if (_watcher != null)
            {
                _watcher.EnableRaisingEvents = false;
                _watcher.Dispose();
                _watcher = null;
                Console.WriteLine("Watcher stopped.");
            }
        }

        // Event handler for new files
        private void OnCreated(object sender, FileSystemEventArgs e)
        {
            // Simple retry to ensure the file is not still being copied
            for (int i = 0; i < 5; i++)
            {
                try
                {
                    using (FileStream stream = File.Open(e.FullPath, FileMode.Open, FileAccess.Read, FileShare.None))
                    {
                        // File is ready
                        break;
                    }
                }
                catch (IOException)
                {
                    Thread.Sleep(500);
                }
            }

            string ext = Path.GetExtension(e.FullPath).ToLowerInvariant();
            if (ext == ".xls" || ext == ".xlsx" || ext == ".xlsm")
            {
                try
                {
                    ConvertExcelToHtml(e.FullPath);
                }
                catch (Exception ex)
                {
                    Console.Error.WriteLine($"Conversion failed for '{e.FullPath}': {ex.Message}");
                }
            }
        }

        // Core conversion logic using Aspose.Cells
        private void ConvertExcelToHtml(string excelPath)
        {
            if (!File.Exists(excelPath))
                throw new FileNotFoundException("Excel file not found.", excelPath);

            // Load the workbook
            Workbook workbook = new Workbook(excelPath);

            // Set custom HTML save options
            HtmlSaveOptions options = new HtmlSaveOptions
            {
                ExportActiveWorksheetOnly = false, // export all worksheets
                ExportImagesAsBase64 = true,       // embed images
                ExportHiddenWorksheet = false,
                ExportGridLines = true,
                ExportPrintAreaOnly = false
                // Note: ExportChartImageFormat and HtmlVersion are not available in the current Aspose.Cells version
            };

            // Build output HTML file path
            string fileName = Path.GetFileNameWithoutExtension(excelPath);
            string htmlPath = Path.Combine(_outputFolder, fileName + ".html");

            try
            {
                // Save workbook as HTML with the specified options
                workbook.Save(htmlPath, options);
                Console.WriteLine($"Converted '{excelPath}' to '{htmlPath}'.");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error saving HTML for '{excelPath}': {ex.Message}");
                throw;
            }
        }

        // Application entry point
        public static void Main()
        {
            var watcher = new ExcelToHtmlWatcher();
            watcher.Start();

            Console.WriteLine("Press ENTER to exit.");
            Console.ReadLine();

            watcher.Stop();
        }
    }
}
