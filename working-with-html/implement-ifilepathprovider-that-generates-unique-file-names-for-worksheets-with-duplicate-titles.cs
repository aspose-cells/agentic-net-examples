// Title: How to implement a custom IFilePathProvider that creates unique, filesystem‑safe filenames for worksheets with duplicate titles using Aspose.Cells for .NET
// AI Prompts: Generate a C# class implementing IFilePathProvider that cleans worksheet titles and adds a numeric suffix when the same title appears more than once. | Write code that ensures the output directory exists, allows a configurable file extension, and returns the full path for each unique worksheet file. | Show how to use the custom provider to iterate through a Workbook and save each worksheet as an individual .xlsx file with Aspose.Cells.
// Common Searches: aspocells example for saving each worksheet as a separate file with unique names | C# generate unique filenames for Excel sheets when exporting with Aspose.Cells | how to clean worksheet titles for Windows file system in Aspose.Cells .NET | save individual worksheets to separate files while avoiding filename collisions in Aspose.Cells | set custom file extension when exporting worksheets using Aspose.Cells
// Tags: custom IFilePathProvider implementation Aspose.Cells | unique worksheet filename generation .NET | sanitize Excel sheet title for file system | export each worksheet as separate file | configurable file extension for worksheet export | handle duplicate sheet titles Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace WorksheetSaverApp
{
    // Simple interface for file‑path providers.
    public interface IFilePathProvider
    {
        string GetFullName(string fileName);
    }

    // Generates unique file paths for worksheets, handling duplicate sheet titles.
    // The example defines an IFilePathProvider and a UniqueWorksheetFilePathProvider that sanitizes worksheet names, tracks occurrences, and generates filesystem‑safe, unique file paths with optional numeric suffixes and configurable extensions. It ensures the target folder exists and demonstrates iterating through a Workbook to save each worksheet as an individual file using Aspose.Cells.
    public class UniqueWorksheetFilePathProvider : IFilePathProvider
    {
        // Tracks usage count of each sanitized sheet name.
        private readonly Dictionary<string, int> _nameCounts = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

        // Destination folder.
        private readonly string _folderPath;

        // Desired file extension (including the dot, e.g., ".xlsx").
        private readonly string _extension;

        public UniqueWorksheetFilePathProvider(string folderPath, string extension = "xlsx")
        {
            _folderPath = folderPath;
            _extension = extension.StartsWith(".") ? extension : "." + extension;
            Directory.CreateDirectory(_folderPath); // Ensure the folder exists.
        }

        // Called by the caller to obtain a unique full file name.
        public string GetFullName(string fileName)
        {
            // Make the sheet name safe for the file system.
            string safeName = MakeValidFileName(fileName);

            // Update occurrence count.
            if (_nameCounts.TryGetValue(safeName, out int count))
            {
                count++;
                _nameCounts[safeName] = count;
            }
            else
            {
                count = 1;
                _nameCounts[safeName] = count;
            }

            // First occurrence uses the plain name; later ones get a numeric suffix.
            string uniqueFileName = count == 1
                ? $"{safeName}{_extension}"
                : $"{safeName}_{count}{_extension}";

            return Path.Combine(_folderPath, uniqueFileName);
        }

        // Replaces invalid file‑system characters with underscores.
        private static string MakeValidFileName(string name)
        {
            char[] invalidChars = Path.GetInvalidFileNameChars();
            foreach (char c in invalidChars)
            {
                name = name.Replace(c.ToString(), "_");
            }
            return name.TrimEnd(' ', '.');
        }
    }

    // Demonstrates saving each worksheet as an individual file.
    public class WorksheetSaverExample
    {
        public void SaveWorksheetsIndividually()
        {
            const string inputPath = "input.xlsx";

            try
            {
                // Verify the input file exists.
                if (!File.Exists(inputPath))
                    throw new FileNotFoundException($"Input file not found: {inputPath}");

                // Load the workbook.
                Workbook workbook = new Workbook(inputPath);

                // Provider that creates unique file names in the "Sheets" directory.
                IFilePathProvider pathProvider = new UniqueWorksheetFilePathProvider("Sheets");

                // Process each worksheet.
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    try
                    {
                        // Create a new workbook containing only the current sheet.
                        Workbook singleSheetWorkbook = new Workbook();
                        // Add a copy of the current sheet by name.
                        singleSheetWorkbook.Worksheets.AddCopy(sheet.Name);
                        // Remove the default empty sheet.
                        singleSheetWorkbook.Worksheets.RemoveAt(0);

                        // Get a unique file path for this sheet.
                        string filePath = pathProvider.GetFullName(sheet.Name);

                        // Save the single‑sheet workbook.
                        singleSheetWorkbook.Save(filePath);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Failed to save sheet \"{sheet.Name}\": {ex.Message}");
                    }
                }
            }
            catch (Exception ex)
            {
                // Log or handle top‑level exceptions.
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Application entry point.
    public class Program
    {
        public static void Main(string[] args)
        {
            try
            {
                WorksheetSaverExample example = new WorksheetSaverExample();
                example.SaveWorksheetsIndividually();
                Console.WriteLine("Worksheet saving completed.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unhandled exception: {ex.Message}");
            }
        }
    }
}
