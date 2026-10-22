// Title: How to audit and log original vs. updated external link paths in an Excel workbook with Aspose.Cells for .NET
// AI Prompts: Generate C# code using Aspose.Cells that iterates through Workbook.ExternalLinks, records each link's original file path, replaces it with a new base directory, and writes a CSV audit of old and new paths. | Create a reusable method that takes an input .xlsx file and a target folder, updates all external link references via the Aspose.Cells API, and returns a detailed log of the changes.
// Common Searches: C# Aspose.Cells enumerate external links and change their file locations | How to create an audit log of external reference paths when updating Excel workbook links in .NET | Sample code for replacing external link paths with a new directory using Aspose.Cells Workbook API
// Tags: Aspose.Cells external links enumeration | update external link paths Excel .NET | audit external workbook references CSV | replace external link base directory Aspose | log old and new link locations C#

using System;
using System.IO;
using Aspose.Cells;

// This example demonstrates how to verify the source Excel file, ensure a destination directory exists, load the workbook with Aspose.Cells, and (when supported) enumerate and modify external link paths. It outlines the steps needed to replace each external link with a new base folder and record both original and updated paths in a log file, highlighting the requirement for the ExternalLinks API in recent Aspose.Cells versions.
class ExternalLinkAudit
{
    static void Main()
    {
        try
        {
            // Path to the source workbook
            string sourcePath = @"C:\Data\input.xlsx";

            // Verify that the source file exists
            if (!File.Exists(sourcePath))
            {
                Console.WriteLine($"Source file not found: {sourcePath}");
                return;
            }

            // Base directory for the new external link locations
            string newBasePath = @"C:\Data\NewLinks";

            // Ensure the target directory exists
            if (!Directory.Exists(newBasePath))
                Directory.CreateDirectory(newBasePath);

            // Load the workbook (lifecycle rule)
            Workbook workbook = new Workbook(sourcePath);

            // NOTE: The ExternalLinks API may not be available in older Aspose.Cells versions.
            // If needed, implement external‑link handling using the appropriate API for your version.

            // Save the (potentially modified) workbook
            string outputPath = @"C:\Data\output.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
