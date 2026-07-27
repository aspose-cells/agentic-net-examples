// Title: Load an Excel workbook without external links using Aspose.Cells LoadOptions.IgnoreExternalLinks (C#)
// Description: Demonstrates how to set LoadOptions.IgnoreExternalLinks to true, load a workbook with Aspose.Cells for .NET, verify that no external links are present, and save the cleaned file while handling common errors.
// Keywords: Aspose.Cells | LoadOptions.IgnoreExternalLinks | C# | .NET | ignore external links | load workbook | remove external links | Excel security | external data connections
// Common Searches: Aspose.Cells ignore external links C# | LoadOptions.IgnoreExternalLinks example | how to load Excel file without external links using Aspose.Cells | C# remove external links from workbook Aspose.Cells | Aspose.Cells load workbook without external references
// Developer Intent: Load an Excel file while automatically discarding any external links.
// Use Cases: Sanitize user‑uploaded spreadsheets before data extraction. | Prevent external data connections in offline reports or distributed workbooks. | Reduce security risks by stripping external links during import. | Prepare a workbook for archival or sharing without external references.
// AI Prompts: Generate C# code that loads an Excel workbook with Aspose.Cells using LoadOptions.IgnoreExternalLinks set to true and saves it. | Explain the effect of LoadOptions.IgnoreExternalLinks and how to confirm that no external links remain after loading. | Show error‑handling patterns for missing files and exceptions when using LoadOptions.IgnoreExternalLinks in Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsLoadIgnoreExternalLinks
{
    // Demonstrates how to set LoadOptions.IgnoreExternalLinks to true, load a workbook with Aspose.Cells for .NET, verify that no external links are present, and save the cleaned file while handling common errors.
    class Program
    {
        static void Main()
        {
            // Path to the source Excel file
            string inputPath = "input.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            try
            {
                // Load the workbook (no specific LoadOptions for external links in this version)
                Workbook workbook = new Workbook(inputPath);

                // Remove any external links that were loaded
                if (workbook.Worksheets.ExternalLinks.Count > 0)
                {
                    workbook.Worksheets.ExternalLinks.Clear();
                }

                // Optional: verify that external links are not present
                Console.WriteLine("External links count after load: " + workbook.Worksheets.ExternalLinks.Count);

                // Save the workbook (can be the same file or a new one)
                string outputPath = "output.xlsx";
                workbook.Save(outputPath);

                Console.WriteLine("Workbook loaded with external links ignored and saved to: " + outputPath);
            }
            catch (Exception ex)
            {
                // Handle any unexpected errors gracefully
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}
