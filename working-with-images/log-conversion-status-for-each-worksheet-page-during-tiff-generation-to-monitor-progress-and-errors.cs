// Title: How to log each worksheet page’s conversion status during multi‑page TIFF export with Aspose.Cells for .NET
// Common Searches: C# log progress of each worksheet page while exporting to multi‑page TIFF using Aspose.Cells | Aspose.Cells IPageSavingCallback example to track page start and end events in TIFF conversion | How to write a custom logger for TIFF page rendering errors in Aspose.Cells .NET | Monitor worksheet to TIFF conversion status per page with Aspose.Cells and C# | Save Excel worksheets as separate TIFF files and log conversion details using Aspose.Cells
// Tags: IPageSavingCallback implementation for TIFF export | C# logging during multi-page TIFF generation | worksheet to TIFF conversion with progress tracking | custom logger for Aspose.Cells image rendering | multi-page TIFF rendering from Excel using SheetRender

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Drawing;

namespace AsposeCellsTiffLogging
{
    // Simple logger that writes to console and optionally to a file
    // The sample demonstrates how to monitor and log the conversion of each worksheet page to a multi‑page TIFF using Aspose.Cells for .NET. It includes a simple Logger class, an IPageSavingCallback implementation that records page start and end events, and code that iterates through all worksheets, configures ImageOrPrintOptions for TIFF output, and renders each sheet with SheetRender while capturing progress and handling errors.
    public class Logger
    {
        private readonly string? _logFilePath;
        private readonly object _lock = new object();

        public Logger(string? logFilePath = null)
        {
            _logFilePath = logFilePath;
            if (!string.IsNullOrEmpty(_logFilePath))
            {
                // Ensure directory exists
                Directory.CreateDirectory(Path.GetDirectoryName(_logFilePath)!);
                // Start fresh
                File.WriteAllText(_logFilePath, $"Log started at {DateTime.Now}{Environment.NewLine}");
            }
        }

        public void Info(string message) => Write($"INFO: {message}");

        public void Error(string message) => Write($"ERROR: {message}");

        private void Write(string message)
        {
            string line = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff} {message}";
            lock (_lock)
            {
                Console.WriteLine(line);
                if (!string.IsNullOrEmpty(_logFilePath))
                {
                    File.AppendAllText(_logFilePath, line + Environment.NewLine);
                }
            }
        }
    }

    // Callback implementation to monitor page saving progress
    public class TiffPageSavingCallback : IPageSavingCallback
    {
        private readonly string _workbookName;
        private readonly string _worksheetName;
        private readonly Logger _logger;

        public TiffPageSavingCallback(string workbookName, string worksheetName, Logger logger)
        {
            _workbookName = workbookName;
            _worksheetName = worksheetName;
            _logger = logger;
        }

        // Called before a page is saved
        public void PageStartSaving(PageStartSavingArgs args)
        {
            try
            {
                _logger.Info($"Workbook '{_workbookName}', Worksheet '{_worksheetName}': Starting to save page {args.PageIndex + 1} of {args.PageCount}.");
                // Ensure the page will be output; you could set args.IsToOutput = false to skip a page
                args.IsToOutput = true;
            }
            catch (Exception ex)
            {
                _logger.Error($"Exception in PageStartSaving for page {args.PageIndex}: {ex.Message}");
                // If an exception occurs, prevent output of this page
                args.IsToOutput = false;
            }
        }

        // Called after a page has been saved
        public void PageEndSaving(PageEndSavingArgs args)
        {
            try
            {
                _logger.Info($"Workbook '{_workbookName}', Worksheet '{_worksheetName}': Finished saving page {args.PageIndex + 1}.");
                // You can control further processing via args.HasMorePages if needed
                args.HasMorePages = true;
            }
            catch (Exception ex)
            {
                _logger.Error($"Exception in PageEndSaving for page {args.PageIndex}: {ex.Message}");
                // Continue processing despite the error
                args.HasMorePages = true;
            }
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            // Expected arguments:
            // args[0] - input workbook path (e.g., "input.xlsx")
            // args[1] - output directory for TIFF files (e.g., "C:\Output")
            // args[2] - optional log file path (e.g., "C:\Output\conversion.log")
            if (args.Length < 2)
            {
                Console.WriteLine("Usage: AsposeCellsTiffLogging.exe <inputWorkbookPath> <outputDirectory> [logFilePath]");
                return;
            }

            string inputPath = args[0];
            string outputDir = args[1];
            string? logPath = args.Length >= 3 ? args[2] : null;

            var logger = new Logger(logPath);

            try
            {
                if (!File.Exists(inputPath))
                {
                    logger.Error($"Input file does not exist: {inputPath}");
                    return;
                }

                Directory.CreateDirectory(outputDir);
                logger.Info($"Loading workbook from '{inputPath}'.");
                Workbook workbook = new Workbook(inputPath);
                string workbookName = Path.GetFileNameWithoutExtension(inputPath);

                // Iterate through each worksheet and render to a multi‑page TIFF
                for (int sheetIndex = 0; sheetIndex < workbook.Worksheets.Count; sheetIndex++)
                {
                    Worksheet sheet = workbook.Worksheets[sheetIndex];
                    string sheetName = sheet.Name;
                    logger.Info($"Processing worksheet {sheetIndex + 1}/{workbook.Worksheets.Count}: '{sheetName}'.");

                    // Configure image options for TIFF output
                    ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
                    {
                        ImageType = ImageType.Tiff,
                        TiffCompression = TiffCompression.CompressionLZW,
                        HorizontalResolution = 300,
                        VerticalResolution = 300,
                        TiffColorDepth = ColorDepth.Format24bpp,
                        PageSavingCallback = new TiffPageSavingCallback(workbookName, sheetName, logger)
                    };

                    // SheetRender handles page splitting based on the worksheet's PageSetup
                    SheetRender renderer = new SheetRender(sheet, imgOptions);

                    // Build output file name: <WorkbookName>_<SheetName>.tiff
                    string safeSheetName = MakeFileSystemSafe(sheetName);
                    string outputPath = Path.Combine(outputDir, $"{workbookName}_{safeSheetName}.tiff");

                    try
                    {
                        logger.Info($"Rendering worksheet '{sheetName}' to TIFF file '{outputPath}'.");
                        // This will generate a multi‑page TIFF where each page corresponds to a printed page
                        renderer.ToTiff(outputPath);
                        logger.Info($"Successfully saved TIFF for worksheet '{sheetName}'.");
                    }
                    catch (Exception renderEx)
                    {
                        logger.Error($"Failed to render worksheet '{sheetName}' to TIFF. Exception: {renderEx.Message}");
                    }
                }

                logger.Info("All worksheets processed.");
            }
            catch (Exception ex)
            {
                logger.Error($"Unexpected error: {ex.Message}");
                logger.Error($"StackTrace: {ex.StackTrace}");
            }
        }

        // Helper to replace characters that are invalid in file names
        private static string MakeFileSystemSafe(string name)
        {
            foreach (char c in Path.GetInvalidFileNameChars())
            {
                name = name.Replace(c, '_');
            }
            return name;
        }
    }
}
