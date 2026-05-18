using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExternalLinkUpdate
{
    public static class ExternalLinkPathUpdater
    {
        public static void UpdateExternalLinkPaths(string inputFile, string outputFile, Func<string, string> pathTransformer)
        {
            if (string.IsNullOrEmpty(inputFile))
                throw new ArgumentException("Input file path must be provided.", nameof(inputFile));
            if (string.IsNullOrEmpty(outputFile))
                throw new ArgumentException("Output file path must be provided.", nameof(outputFile));
            if (pathTransformer == null)
                throw new ArgumentNullException(nameof(pathTransformer));

            // Ensure the input file exists; if not, create an empty workbook.
            if (!File.Exists(inputFile))
            {
                var wb = new Workbook();
                wb.Save(inputFile);
            }

            using (Workbook workbook = new Workbook(inputFile))
            {
                ExternalLinkCollection externalLinks = workbook.Worksheets.ExternalLinks;

                for (int i = 0; i < externalLinks.Count; i++)
                {
                    ExternalLink link = externalLinks[i];
                    string originalPath = link.DataSource;

                    try
                    {
                        string newPath = pathTransformer(originalPath);
                        link.DataSource = newPath;
                        Console.WriteLine($"External link #{i} updated from '{originalPath}' to '{newPath}'.");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Failed to update external link #{i} (original path: '{originalPath}'): {ex.Message}");
                    }
                }

                try
                {
                    workbook.Save(outputFile);
                    Console.WriteLine($"Workbook saved successfully to '{outputFile}'.");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error saving workbook to '{outputFile}': {ex.Message}");
                }
            }
        }

        public static void RunExample()
        {
            string sourcePath = Path.Combine(Environment.CurrentDirectory, "SampleWithLinks.xlsx");
            string destinationPath = Path.Combine(Environment.CurrentDirectory, "SampleWithLinks_Updated.xlsx");

            Func<string, string> transformer = currentPath =>
            {
                if (string.IsNullOrEmpty(currentPath))
                    throw new InvalidOperationException("Current path is null or empty.");

                const string oldPrefix = @"\\Server\Share\";
                const string newPrefix = @"D:\Data\";

                if (!currentPath.StartsWith(oldPrefix, StringComparison.OrdinalIgnoreCase))
                    return currentPath;

                string relativePart = currentPath.Substring(oldPrefix.Length);
                return Path.Combine(newPrefix, relativePart);
            };

            UpdateExternalLinkPaths(sourcePath, destinationPath, transformer);
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            ExternalLinkPathUpdater.RunExample();
        }
    }
}