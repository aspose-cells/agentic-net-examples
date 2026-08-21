// Title: C# – Validate ExternalLink DataSource File Exists Before Updating with Aspose.Cells
// Description: Loads an Excel workbook, verifies that a specified external workbook file is present, updates every ExternalLink.DataSource to the new path, and saves the workbook, with robust handling for missing files.
// Keywords: Aspose.Cells ExternalLink | C# validate external link | check external workbook existence | update DataSource Aspose.Cells | Excel external link validation | file existence check C# | Aspose.Cells workbook external references
// Common Searches: How to check if external Excel file exists before setting ExternalLink.DataSource in Aspose.Cells C# | Aspose.Cells C# update external link path with validation | Validate external link source file in workbook using Aspose.Cells | C# Aspose.Cells external link file existence check | Prevent missing external file errors Aspose.Cells
// Developer Intent: Confirm that the new external link points to an existing file before modifying the workbook’s ExternalLink.DataSource.
// Use Cases: Avoid runtime failures when redirecting external references by ensuring the target file is available. | Batch‑process multiple workbooks to point their external links to a verified data source. | Integrate a safety check into deployment scripts that update external links only after confirming file presence.
// AI Prompts: Generate C# code using Aspose.Cells that updates all ExternalLink.DataSource values only after confirming each target file exists. | Create detailed error‑handling and logging for missing external Excel files when changing ExternalLink.DataSource in a workbook. | Refactor the sample to support multiple new external paths, log validation results, and return a summary of successful updates.

using System;
using System.IO;
using Aspose.Cells;

// Loads an Excel workbook, verifies that a specified external workbook file is present, updates every ExternalLink.DataSource to the new path, and saves the workbook, with robust handling for missing files.
class ValidateExternalLink
{
    public static void Run()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook (lifecycle rule: load)
            Workbook workbook = new Workbook(inputPath);

            // Define the new external link path
            string newExternalPath = @"C:\Data\external.xlsx";

            // Validate that the external file exists before applying the change
            if (!File.Exists(newExternalPath))
            {
                Console.WriteLine($"The specified external file does not exist: {newExternalPath}");
                return;
            }

            // Update the DataSource of each external link (feature rule: ExternalLink.DataSource)
            foreach (ExternalLink link in workbook.Worksheets.ExternalLinks)
            {
                link.DataSource = newExternalPath;
            }

            // Save the workbook with the updated external link (lifecycle rule: save)
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }

    // Entry point for the application
    static void Main(string[] args)
    {
        Run();
    }
}
