using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Saving;

namespace AsposeCellsDemo
{
    public class WorkbookExportHandler
    {
        // Exports an existing workbook to a file on disk.
        public void ExportWorkbook(string outputPath)
        {
            try
            {
                const string inputPath = "input.xlsx";

                // Verify that the source workbook exists.
                if (!File.Exists(inputPath))
                    throw new FileNotFoundException($"The workbook file '{inputPath}' was not found.");

                // Load the workbook from the file.
                var workbook = new Workbook(inputPath);

                // Use OoxmlSaveOptions for .xlsx output.
                var saveOptions = new OoxmlSaveOptions(SaveFormat.Xlsx);

                // Save the workbook to the specified output path.
                workbook.Save(outputPath, saveOptions);
            }
            catch (Exception ex)
            {
                // Wrap and rethrow to allow the caller to handle the error.
                throw new InvalidOperationException("Failed to export workbook.", ex);
            }
        }
    }

    public class Program
    {
        // Entry point required for console application.
        public static void Main(string[] args)
        {
            try
            {
                string outputPath = "output.xlsx";

                // Ensure the output directory exists.
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                var handler = new WorkbookExportHandler();
                handler.ExportWorkbook(outputPath);

                Console.WriteLine($"Workbook exported successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}