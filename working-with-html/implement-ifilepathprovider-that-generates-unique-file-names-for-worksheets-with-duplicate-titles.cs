// Title: C# example: Export each worksheet to a separate .xlsx file with unique filenames using a custom IFilePathProvider in Aspose.Cells
// Common Searches: how to export each sheet of an Excel workbook to its own file using Aspose.Cells C# | c# generate unique filenames for worksheets with duplicate names when saving with Aspose.Cells | implement IFilePathProvider to avoid filename collisions in Aspose.Cells worksheet export | sanitize worksheet names for file system when exporting with Aspose.Cells | save individual worksheets as separate workbooks Aspose.Cells C# example
// Tags: Aspose.Cells custom IFilePathProvider implementation | unique worksheet filename generation C# | export worksheets to individual .xlsx files | sanitize worksheet name for file path Aspose.Cells | prevent duplicate sheet title file collisions | WorksheetExporter class Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsWorksheetExport
{
    // Provides a full file path for a worksheet, ensuring uniqueness.
    public interface IFilePathProvider
    {
        string GetFilePath(Worksheet worksheet, int index);
    }

    // Generates unique file names inside a base folder.
    // Demonstrates how to implement a custom IFilePathProvider that sanitizes worksheet names, ensures the output directory exists, and generates unique .xlsx file paths to export each worksheet of a source workbook as a separate workbook using Aspose.Cells for .NET.
    public class UniqueWorksheetFilePathProvider : IFilePathProvider
    {
        private readonly string _baseFolder;
        private readonly HashSet<string> _usedNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        public UniqueWorksheetFilePathProvider(string baseFolder)
        {
            _baseFolder = baseFolder ?? throw new ArgumentNullException(nameof(baseFolder));
            Directory.CreateDirectory(_baseFolder); // Ensure the folder exists.
        }

        public string GetFilePath(Worksheet worksheet, int index)
        {
            // Start with the worksheet name.
            string baseName = worksheet.Name ?? $"Sheet{index + 1}";

            // Replace illegal file name characters.
            foreach (char c in Path.GetInvalidFileNameChars())
            {
                baseName = baseName.Replace(c.ToString(), "_");
            }

            // Initial candidate file name.
            string fileName = $"{baseName}.xlsx";

            // Ensure uniqueness across this session and on disk.
            int suffix = 1;
            while (_usedNames.Contains(fileName) || File.Exists(Path.Combine(_baseFolder, fileName)))
            {
                fileName = $"{baseName}_{suffix}.xlsx";
                suffix++;
            }

            _usedNames.Add(fileName);
            return Path.Combine(_baseFolder, fileName);
        }
    }

    // Handles exporting each worksheet to its own workbook file.
    public class WorksheetExporter
    {
        private readonly IFilePathProvider _pathProvider;

        public WorksheetExporter(IFilePathProvider pathProvider)
        {
            _pathProvider = pathProvider ?? throw new ArgumentNullException(nameof(pathProvider));
        }

        public void ExportWorksheets(string sourceWorkbookPath)
        {
            if (string.IsNullOrWhiteSpace(sourceWorkbookPath))
                throw new ArgumentException("Source workbook path must be provided.", nameof(sourceWorkbookPath));

            if (!File.Exists(sourceWorkbookPath))
                throw new FileNotFoundException("Source workbook not found.", sourceWorkbookPath);

            try
            {
                // Load the source workbook.
                Workbook sourceWorkbook = new Workbook(sourceWorkbookPath);

                for (int i = 0; i < sourceWorkbook.Worksheets.Count; i++)
                {
                    Worksheet ws = sourceWorkbook.Worksheets[i];

                    try
                    {
                        // Create a new workbook containing only this worksheet.
                        Workbook singleSheetWorkbook = new Workbook(); // contains a default empty sheet
                        // Add a copy of the current worksheet by name (Aspose.Cells expects sheet name).
                        singleSheetWorkbook.Worksheets.AddCopy(ws.Name);
                        // Remove the default empty sheet created by the constructor.
                        singleSheetWorkbook.Worksheets.RemoveAt(0);

                        // Determine a unique output path.
                        string outputPath = _pathProvider.GetFilePath(ws, i);

                        // Save the single‑sheet workbook.
                        singleSheetWorkbook.Save(outputPath, SaveFormat.Xlsx);
                    }
                    catch (Exception ex)
                    {
                        Console.Error.WriteLine($"Failed to export worksheet '{ws.Name}': {ex.Message}");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error during export: {ex.Message}");
                throw;
            }
        }
    }

    // Entry point for the console application.
    public static class Program
    {
        public static void Main(string[] args)
        {
            // Expect two arguments: source workbook path and output folder.
            if (args.Length < 2)
            {
                Console.WriteLine("Usage: AsposeCellsWorksheetExport <sourceWorkbookPath> <outputFolder>");
                return;
            }

            string sourcePath = args[0];
            string outputFolder = args[1];

            try
            {
                IFilePathProvider provider = new UniqueWorksheetFilePathProvider(outputFolder);
                WorksheetExporter exporter = new WorksheetExporter(provider);
                exporter.ExportWorksheets(sourcePath);
                Console.WriteLine("Export completed successfully.");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Fatal error: {ex.Message}");
            }
        }
    }
}
