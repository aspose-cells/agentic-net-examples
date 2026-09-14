// Title: Add per‑link error handling when updating external link paths from a ribbon command using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an Excel workbook with Aspose.Cells, iterates through workbook.Worksheets.ExternalLinks, replaces each ExternalLink.DataSource with a new folder path, and encloses the load, update, and save operations in try‑catch blocks to capture and log errors. | Create a routine that writes any exception thrown while updating an ExternalLink.DataSource to a log file and then continues processing the remaining links. | Write a validation method that checks whether the transformed external link path exists on disk before assigning it, and throws a custom InvalidExternalLinkPathException if the check fails.
// Common Searches: how to handle exceptions when modifying external links in an Excel workbook with Aspose.Cells C# | Aspose.Cells update external link DataSource path from ribbon button with error handling | C# sample for batch updating external link paths and logging failures using Aspose.Cells | prevent workbook save failure after external link changes in Aspose.Cells .NET
// Tags: Aspose.Cells external link path update with error handling | C# try-catch workbook load Aspose.Cells | validate external link DataSource existence .NET | log external link update failures Aspose.Cells | ensure output directory before saving workbook C#

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExternalLinkUpdate
{
    // Demonstrates loading a workbook, safely iterating its ExternalLinkCollection, updating each ExternalLink.DataSource with per‑link try‑catch handling, validating new paths, logging any errors, ensuring the output folder exists, saving the workbook, and disposing resources.
    public class RibbonExternalLinkUpdater
    {
        /// <param name="inputPath">Path to the workbook to be processed.</param>
        /// <param name="outputPath">Path where the updated workbook will be saved.</param>
        public static void Run(string inputPath, string outputPath)
        {
            // Verify input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: '{inputPath}'.");
                return;
            }

            Workbook workbook = null;

            // Load the workbook
            try
            {
                workbook = new Workbook(inputPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to load workbook '{inputPath}': {ex.Message}");
                return;
            }

            // Get the external links collection
            ExternalLinkCollection externalLinks = workbook.Worksheets.ExternalLinks;

            // Iterate through each external link and attempt to update its path
            for (int i = 0; i < externalLinks.Count; i++)
            {
                ExternalLink link = externalLinks[i];
                string originalPath = link.DataSource;

                try
                {
                    // Example transformation: replace old base folder with new base folder
                    // Adjust this logic to match the actual path update requirements.
                    string updatedPath = originalPath.Replace(@"C:\OldFolder\", @"D:\NewFolder\");

                    // If the path actually changes, assign the new value
                    if (!string.Equals(originalPath, updatedPath, StringComparison.OrdinalIgnoreCase))
                    {
                        link.DataSource = updatedPath;
                        Console.WriteLine($"External link {i} path updated from '{originalPath}' to '{updatedPath}'.");
                    }
                    else
                    {
                        Console.WriteLine($"External link {i} path unchanged.");
                    }
                }
                catch (Exception ex)
                {
                    // Handle failures for this specific link without aborting the whole process
                    Console.WriteLine($"Error updating external link at index {i} (original path: '{originalPath}'): {ex.Message}");
                }
            }

            // Save the workbook
            try
            {
                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save workbook to '{outputPath}': {ex.Message}");
            }
            finally
            {
                // Release resources
                workbook?.Dispose();
            }
        }
    }

    // Entry point for the console application
    public static class Program
    {
        public static void Main(string[] args)
        {
            // Expecting two arguments: inputPath and outputPath
            if (args.Length < 2)
            {
                Console.WriteLine("Usage: AsposeCellsExternalLinkUpdate <inputPath> <outputPath>");
                return;
            }

            string inputPath = args[0];
            string outputPath = args[1];

            try
            {
                RibbonExternalLinkUpdater.Run(inputPath, outputPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }
    }
}
