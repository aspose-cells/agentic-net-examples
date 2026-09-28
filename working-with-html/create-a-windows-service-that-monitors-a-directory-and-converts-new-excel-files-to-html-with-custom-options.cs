// Title: Build a C# Windows service that watches a folder and converts incoming .xlsx files to HTML5 with Aspose.Cells using custom HtmlSaveOptions
// AI Prompts: Write C# code for a Windows service that uses FileSystemWatcher to detect new .xlsx files, waits until the file is fully written, and then converts each workbook to HTML5 with Aspose.Cells, embedding images as Base64 and preserving gridlines. | Show how to configure Aspose.Cells HtmlSaveOptions to export only the active worksheet, set UTF‑8 encoding, enable HTML5 output, and embed images as Base64 when saving a workbook as HTML. | Implement start and stop methods for a console‑based Windows service that creates the output directory, logs conversion results, and handles errors during Excel‑to‑HTML conversion.
// Common Searches: C# example of a Windows service that monitors a directory and automatically converts new .xlsx files to HTML with Aspose.Cells | Aspose.Cells HtmlSaveOptions settings for embedding images as Base64 and exporting gridlines | How to ensure a newly created Excel file is completely written before loading it with Aspose.Cells in C# | Convert only the first worksheet of an Excel workbook to HTML5 using Aspose.Cells | Create output folder and log conversion status in a C# folder‑watcher service
// Tags: directory monitor excel-to-html conversion | Aspose.Cells HtmlSaveOptions html5 configuration | base64 image embedding in html from excel | detect file write completion c# | windows service folder monitoring

using System;
using System.IO;
using System.Text;
using System.Threading;
using Aspose.Cells;

namespace ExcelToHtmlService
{
    // Simple console application that monitors a folder and converts new Excel files to HTML
    // A C# Windows service that watches C:\WatchFolder for new .xlsx files, waits for each file to be fully written, loads the workbook with Aspose.Cells, and saves it as an HTML5 file in C:\HtmlOutput using HtmlSaveOptions configured for UTF‑8 encoding, Base64‑embedded images, gridlines, and active‑worksheet‑only export. Includes start/stop lifecycle methods, output folder creation, and robust error handling.
    public class ExcelConversionService
    {
        private FileSystemWatcher? _watcher;
        private readonly string _inputFolder = @"C:\WatchFolder";
        private readonly string _outputFolder = @"C:\HtmlOutput";

        public void Start()
        {
            try
            {
                // Ensure output directory exists
                Directory.CreateDirectory(_outputFolder);

                // Initialize FileSystemWatcher
                _watcher = new FileSystemWatcher
                {
                    Path = _inputFolder,
                    Filter = "*.xlsx",
                    NotifyFilter = NotifyFilters.FileName | NotifyFilters.CreationTime,
                    EnableRaisingEvents = true,
                    IncludeSubdirectories = false
                };

                // Subscribe to created event
                _watcher.Created += OnNewExcelFile;

                Console.WriteLine($"{DateTime.Now}: ExcelConversionService started. Watching {_inputFolder}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"{DateTime.Now}: Failed to start service – {ex.Message}");
            }
        }

        public void Stop()
        {
            try
            {
                if (_watcher != null)
                {
                    _watcher.EnableRaisingEvents = false;
                    _watcher.Created -= OnNewExcelFile;
                    _watcher.Dispose();
                    _watcher = null;
                }

                Console.WriteLine($"{DateTime.Now}: ExcelConversionService stopped.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"{DateTime.Now}: Error stopping service – {ex.Message}");
            }
        }

        private void OnNewExcelFile(object sender, FileSystemEventArgs e)
        {
            // Wait for the file to be fully written
            try
            {
                const int maxAttempts = 5;
                const int delayMs = 500;

                for (int i = 0; i < maxAttempts; i++)
                {
                    try
                    {
                        using (FileStream stream = File.Open(e.FullPath, FileMode.Open, FileAccess.Read, FileShare.None))
                        {
                            // File opened exclusively – ready for processing
                            break;
                        }
                    }
                    catch (IOException)
                    {
                        Thread.Sleep(delayMs);
                    }
                }

                ConvertExcelToHtml(e.FullPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"{DateTime.Now}: Error processing file {e.Name}: {ex.Message}");
            }
        }

        private void ConvertExcelToHtml(string excelPath)
        {
            try
            {
                if (!File.Exists(excelPath))
                {
                    Console.WriteLine($"{DateTime.Now}: File not found – {excelPath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(excelPath);

                // Configure HTML save options
                HtmlSaveOptions saveOptions = new HtmlSaveOptions
                {
                    ExportActiveWorksheetOnly = true,   // Export only the first worksheet
                    ExportImagesAsBase64 = true,        // Embed images directly
                    ExportGridLines = true,             // Show grid lines
                    HtmlVersion = HtmlVersion.Html5,
                    Encoding = Encoding.UTF8
                };

                // Determine output file name
                string fileNameWithoutExt = Path.GetFileNameWithoutExtension(excelPath);
                string htmlPath = Path.Combine(_outputFolder, fileNameWithoutExt + ".html");

                // Save as HTML
                workbook.Save(htmlPath, saveOptions);

                Console.WriteLine($"{DateTime.Now}: Converted '{excelPath}' to HTML at '{htmlPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"{DateTime.Now}: Conversion failed for '{excelPath}' – {ex.Message}");
            }
        }

        // Main entry point for the console application
        public static void Main()
        {
            ExcelConversionService service = new ExcelConversionService();
            service.Start();

            Console.WriteLine("Press Enter to stop the service...");
            Console.ReadLine();

            service.Stop();
        }
    }
}
