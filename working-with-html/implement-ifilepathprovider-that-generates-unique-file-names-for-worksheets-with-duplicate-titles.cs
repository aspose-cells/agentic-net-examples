// Title: Create a custom IFilePathProvider that generates unique, file‑system‑safe names for worksheets with duplicate titles in Aspose.Cells for .NET
// Common Searches: how to implement IFilePathProvider in Aspose.Cells to avoid filename collisions for worksheets | Aspose.Cells export each worksheet to a separate file with unique name when sheet titles repeat | C# generate safe file names from worksheet titles using Aspose.Cells low‑code API | unique filename strategy for duplicate sheet names in Aspose.Cells .NET | custom file path provider example for splitting an Aspose.Cells workbook into individual sheets
// Tags: custom IFilePathProvider Aspose.Cells | unique worksheet filename generation | duplicate sheet title handling .NET | export worksheet to individual workbook | file system safe sheet names | worksheet cloning with unique file paths

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsFilePathProviderDemo
{
    // The interface is part of Aspose.Cells.LowCode namespace in the real SDK.
    // For demonstration we define it here with the expected signature.
    public interface IFilePathProvider
    {
        /// <param name="worksheetTitle">The title (Name property) of the worksheet.</param>
        /// <returns>A unique file name (with extension) for the worksheet.</returns>
        string GetFilePath(string worksheetTitle);
    }

    // The UniqueWorksheetFilePathProvider class implements Aspose.Cells' IFilePathProvider to produce Windows‑safe file names for worksheets. It normalizes titles, tracks occurrence counts, appends incremental suffixes for duplicates, and returns a full path that combines the output directory, a unique file name, and the chosen extension, enabling reliable export of each worksheet to its own workbook.
    public sealed class UniqueWorksheetFilePathProvider : IFilePathProvider
    {
        private readonly string _outputDirectory;
        private readonly string _extension;
        private readonly Dictionary<string, int> _titleCounters = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

        /// <param name="outputDirectory">Directory where files will be saved. Must exist.</param>
        /// <param name="extension">File extension without leading dot (e.g., "xlsx").</param>
        public UniqueWorksheetFilePathProvider(string outputDirectory, string extension = "xlsx")
        {
            if (string.IsNullOrWhiteSpace(outputDirectory))
                throw new ArgumentException("Output directory cannot be null or whitespace.", nameof(outputDirectory));

            if (string.IsNullOrWhiteSpace(extension))
                throw new ArgumentException("Extension cannot be null or whitespace.", nameof(extension));

            _outputDirectory = outputDirectory.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
            _extension = extension.StartsWith(".") ? extension.Substring(1) : extension;
        }

        /// <inheritdoc/>
        public string GetFilePath(string worksheetTitle)
        {
            if (worksheetTitle == null)
                throw new ArgumentNullException(nameof(worksheetTitle));

            // Normalise the title to a safe file name (remove invalid chars)
            string safeTitle = MakeFileSystemSafe(worksheetTitle);

            // Determine the next suffix for this title
            int count = 0;
            if (_titleCounters.TryGetValue(safeTitle, out int existing))
            {
                count = existing + 1;
                _titleCounters[safeTitle] = count;
            }
            else
            {
                _titleCounters[safeTitle] = 0;
            }

            string fileName = count == 0
                ? $"{safeTitle}.{_extension}"
                : $"{safeTitle}_{count}.{_extension}";

            return Path.Combine(_outputDirectory, fileName);
        }

        // Removes characters that are illegal in Windows file names.
        private static string MakeFileSystemSafe(string name)
        {
            char[] invalid = Path.GetInvalidFileNameChars();
            foreach (char c in invalid)
            {
                name = name.Replace(c.ToString(), "_");
            }
            // Trim trailing spaces or periods which are also illegal.
            return name.TrimEnd(' ', '.');
        }
    }

    class Program
    {
        static void Main()
        {
            try
            {
                // -----------------------------------------------------------------
                // 1. Prepare a workbook with duplicate worksheet titles.
                // -----------------------------------------------------------------
                Workbook sourceWorkbook = new Workbook();

                // First worksheet – default name "Sheet1"
                Worksheet ws1 = sourceWorkbook.Worksheets[0];
                ws1.Name = "Report";
                ws1.Cells["A1"].PutValue("Data for first Report sheet");

                // Add a second worksheet with the same title "Report"
                Worksheet ws2 = sourceWorkbook.Worksheets.Add("Report");
                ws2.Cells["A1"].PutValue("Data for second Report sheet");

                // Add a third worksheet with a distinct title
                Worksheet ws3 = sourceWorkbook.Worksheets.Add("Summary");
                ws3.Cells["A1"].PutValue("Summary data");

                // Add a fourth worksheet that also collides with "Summary"
                Worksheet ws4 = sourceWorkbook.Worksheets.Add("Summary");
                ws4.Cells["A1"].PutValue("Another summary sheet");

                // -----------------------------------------------------------------
                // 2. Create the file‑path provider.
                // -----------------------------------------------------------------
                string outputDir = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
                    "AsposeCellsDemo");

                // Ensure the directory exists.
                Directory.CreateDirectory(outputDir);

                IFilePathProvider pathProvider = new UniqueWorksheetFilePathProvider(outputDir, "xlsx");

                // -----------------------------------------------------------------
                // 3. Export each worksheet to its own workbook using the provider.
                // -----------------------------------------------------------------
                for (int i = 0; i < sourceWorkbook.Worksheets.Count; i++)
                {
                    Worksheet srcSheet = sourceWorkbook.Worksheets[i];
                    string targetPath = pathProvider.GetFilePath(srcSheet.Name);

                    // Create a new workbook that will contain only this sheet.
                    using (Workbook singleSheetWorkbook = new Workbook())
                    {
                        // Remove the default sheet that Aspose.Cells creates.
                        singleSheetWorkbook.Worksheets.Clear();

                        // Clone the source worksheet into the new workbook.
                        // AddCopy returns the index of the newly added sheet.
                        int newIndex = singleSheetWorkbook.Worksheets.AddCopy(srcSheet.Name);
                        Worksheet cloned = singleSheetWorkbook.Worksheets[newIndex];

                        // Preserve the original sheet name.
                        cloned.Name = srcSheet.Name;

                        try
                        {
                            singleSheetWorkbook.Save(targetPath, SaveFormat.Xlsx);
                            Console.WriteLine($"Saved worksheet \"{srcSheet.Name}\" to \"{targetPath}\"");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"Failed to save worksheet \"{srcSheet.Name}\": {ex.Message}");
                        }
                    }
                }

                // Clean up the original workbook.
                sourceWorkbook.Dispose();

                Console.WriteLine("All worksheets have been exported.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }
    }
}
