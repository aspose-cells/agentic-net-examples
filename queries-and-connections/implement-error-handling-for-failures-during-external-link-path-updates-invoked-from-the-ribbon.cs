using System;
using Aspose.Cells;

namespace AsposeCellsExternalLinkHandling
{
    public class ExternalLinkPathUpdater
    {
        /// <summary>
        /// Loads a workbook, updates the DataSource of each external link using the supplied
        /// path transformation function, and saves the workbook.
        /// All operations are wrapped with error handling to capture and report failures.
        /// </summary>
        /// <param name="inputFile">Full path of the workbook to process.</param>
        /// <param name="outputFile">Full path where the updated workbook will be saved.</param>
        /// <param name="transformPath">
        /// A function that receives the current DataSource string and returns the new path.
        /// </param>
        public void UpdateExternalLinkPaths(string inputFile, string outputFile, Func<string, string> transformPath)
        {
            if (string.IsNullOrEmpty(inputFile))
                throw new ArgumentException("Input file path must be provided.", nameof(inputFile));

            if (string.IsNullOrEmpty(outputFile))
                throw new ArgumentException("Output file path must be provided.", nameof(outputFile));

            if (transformPath == null)
                throw new ArgumentNullException(nameof(transformPath));

            Workbook workbook = null;

            try
            {
                // Load the workbook
                workbook = new Workbook(inputFile);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error loading workbook '{inputFile}': {ex.Message}");
                return;
            }

            // Get the external links collection
            ExternalLinkCollection externalLinks = workbook.Worksheets.ExternalLinks;

            // Iterate through each external link and attempt to update its DataSource
            for (int i = 0; i < externalLinks.Count; i++)
            {
                try
                {
                    ExternalLink link = externalLinks[i];
                    string originalPath = link.DataSource;

                    // Transform the path using the user‑provided delegate
                    string newPath = transformPath(originalPath);

                    // Only assign if the path actually changed
                    if (!string.Equals(originalPath, newPath, StringComparison.OrdinalIgnoreCase))
                    {
                        link.DataSource = newPath;
                        Console.WriteLine($"External link #{i} updated: '{originalPath}' => '{newPath}'");
                    }
                }
                catch (Exception ex)
                {
                    // Capture any failure for the specific link but continue processing others
                    Console.WriteLine($"Error updating external link at index {i}: {ex.Message}");
                }
            }

            try
            {
                // Save the modified workbook
                workbook.Save(outputFile);
                Console.WriteLine($"Workbook saved successfully to '{outputFile}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error saving workbook to '{outputFile}': {ex.Message}");
            }
            finally
            {
                // Ensure resources are released
                workbook?.Dispose();
            }
        }
    }

    // Example usage (could be wired to a ribbon button)
    class Program
    {
        static void Main()
        {
            var updater = new ExternalLinkPathUpdater();

            // Example transformation: replace an old SharePoint URL with a new one
            Func<string, string> pathTransformer = original =>
                original.Replace(
                    @"https://oldsharepoint.com/Docs/",
                    @"https://newsharepoint.com/Shared/");

            updater.UpdateExternalLinkPaths(
                inputFile: @"C:\Temp\SourceWorkbook.xlsx",
                outputFile: @"C:\Temp\SourceWorkbook_Updated.xlsx",
                transformPath: pathTransformer);
        }
    }
}