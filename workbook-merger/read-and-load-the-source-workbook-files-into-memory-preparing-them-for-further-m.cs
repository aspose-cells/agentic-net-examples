using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsDemo
{
    public static class WorkbookLoader
    {
        /// <summary>
        /// Loads multiple Excel workbooks from the specified file paths into memory.
        /// </summary>
        /// <param name="filePaths">Array of full file paths to the source workbooks.</param>
        /// <returns>List of loaded Workbook objects ready for further manipulation.</returns>
        public static List<Workbook> LoadWorkbooks(string[] filePaths)
        {
            var workbooks = new List<Workbook>();

            foreach (string path in filePaths)
            {
                if (!File.Exists(path))
                {
                    throw new FileNotFoundException($"The file '{path}' was not found.");
                }

                var wb = new Workbook(path);
                workbooks.Add(wb);
            }

            return workbooks;
        }

        /// <summary>
        /// Example usage of LoadWorkbooks method.
        /// </summary>
        public static void Example()
        {
            // Create temporary workbooks for demonstration
            string tempDir = Path.Combine(Path.GetTempPath(), "AsposeCellsDemo");
            Directory.CreateDirectory(tempDir);

            string[] sourceFiles = new string[3];
            for (int i = 0; i < sourceFiles.Length; i++)
            {
                string filePath = Path.Combine(tempDir, $"Source{i + 1}.xlsx");
                var wb = new Workbook();
                wb.Worksheets[0].Name = $"Sheet{i + 1}";
                wb.Save(filePath);
                wb.Dispose();
                sourceFiles[i] = filePath;
            }

            // Load the created workbooks
            List<Workbook> loadedWorkbooks = LoadWorkbooks(sourceFiles);

            for (int i = 0; i < loadedWorkbooks.Count; i++)
            {
                Console.WriteLine($"Workbook {i + 1}: {loadedWorkbooks[i].Worksheets.Count} worksheets loaded.");
            }

            // Clean up
            foreach (var wb in loadedWorkbooks)
            {
                wb.Dispose();
            }

            // Optionally delete temporary files
            foreach (var file in sourceFiles)
            {
                File.Delete(file);
            }
            Directory.Delete(tempDir);
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            WorkbookLoader.Example();
        }
    }
}