// Title: Parallel conversion of Excel worksheets to PNG images using Aspose.Cells and Task Parallel Library in C#
// Common Searches: how to export each worksheet of an Excel file to separate PNG files using Aspose.Cells in C# | using Parallel.ForEach to speed up worksheet image rendering with Aspose.Cells | C# code sample for multi‑core PNG conversion of Excel sheets with Aspose.Cells | best practices for thread‑safe worksheet list when rendering Excel to PNG in .NET | set max degree of parallelism for Aspose.Cells SheetRender PNG export
// Tags: parallel worksheet PNG rendering with Aspose.Cells | Task Parallel Library Excel to PNG conversion | SheetRender multi‑core image export | ImageOrPrintOptions PNG configuration Aspose.Cells | thread‑safe worksheet enumeration for parallel rendering

using System;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace WorksheetToPngParallel
{
    // The example validates input parameters, loads an Excel workbook with Aspose.Cells, creates a thread‑safe list of worksheets, and uses Parallel.ForEach (max degree set to the processor count) to render each sheet to PNG files. ImageOrPrintOptions are set for PNG format with OnePagePerSheet=true, and file names are sanitized for the file system. Errors are handled per worksheet to keep processing alive, and progress is logged with thread IDs.
    class Program
    {
        static void Main(string[] args)
        {
            // Validate arguments
            if (args.Length < 2)
            {
                Console.WriteLine("Usage: WorksheetToPngParallel.exe <inputWorkbookPath> <outputFolder>");
                return;
            }

            string inputPath = args[0];
            string outputFolder = args[1];

            if (!File.Exists(inputPath))
            {
                Console.Error.WriteLine($"Error: Input file \"{inputPath}\" does not exist.");
                return;
            }

            try
            {
                // Ensure output directory exists
                if (!Directory.Exists(outputFolder))
                {
                    Directory.CreateDirectory(outputFolder);
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Create a list of worksheets to avoid collection modification issues during parallel execution
                var worksheets = workbook.Worksheets.Cast<Worksheet>().ToList();

                // Configure parallel options (use all logical processors)
                ParallelOptions parallelOptions = new ParallelOptions
                {
                    MaxDegreeOfParallelism = Environment.ProcessorCount
                };

                // Parallel processing of each worksheet
                Parallel.ForEach(worksheets, parallelOptions, sheet =>
                {
                    try
                    {
                        string sheetName = sheet.Name;

                        // Set image options: PNG format, one page per sheet (entire sheet in a single image)
                        ImageOrPrintOptions options = new ImageOrPrintOptions
                        {
                            ImageType = Aspose.Cells.Drawing.ImageType.Png,
                            OnePagePerSheet = true
                        };

                        // Create a SheetRender for the current worksheet
                        SheetRender renderer = new SheetRender(sheet, options);
                        int pageCount = renderer.PageCount; // Usually 1 when OnePagePerSheet = true

                        for (int pageIndex = 0; pageIndex < pageCount; pageIndex++)
                        {
                            // Build a unique file name for each page
                            string safeSheetName = MakeFileSystemSafe(sheetName);
                            string outputFile = Path.Combine(
                                outputFolder,
                                $"{Path.GetFileNameWithoutExtension(inputPath)}_{safeSheetName}_Page{pageIndex + 1}.png");

                            // Render the page to a PNG file
                            renderer.ToImage(pageIndex, outputFile);
                        }

                        Console.WriteLine($"[Thread {Thread.CurrentThread.ManagedThreadId}] Rendered sheet \"{sheetName}\" ({pageCount} page(s)).");
                    }
                    catch (Exception ex)
                    {
                        // Log any errors for the specific worksheet but continue processing others
                        Console.Error.WriteLine($"[Thread {Thread.CurrentThread.ManagedThreadId}] Error rendering sheet \"{sheet.Name}\": {ex.Message}");
                    }
                });

                Console.WriteLine("All worksheets have been processed successfully.");
            }
            catch (Exception ex)
            {
                // Catch any unexpected errors that occur outside the parallel loop
                Console.Error.WriteLine($"Fatal error: {ex.Message}");
            }
        }

        private static string MakeFileSystemSafe(string name)
        {
            var invalidChars = Path.GetInvalidFileNameChars();
            foreach (char c in invalidChars)
            {
                name = name.Replace(c, '_');
            }
            return name;
        }
    }
}
