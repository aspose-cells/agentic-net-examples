// Title: Copy a Range with Formulas and External Links to a New Workbook using Aspose.Cells for .NET
// Description: Loads a source workbook, creates an empty destination workbook, defines matching source and destination ranges, and copies the range with PasteOptions set to include all content while preserving external links. After the copy, iterates through the destination workbook's ExternalLinks collection and updates each link's DataSource to the new file name before saving.
// Keywords: Aspose.Cells copy range C# | external links update Aspose.Cells | PasteOptions IgnoreLinksToOriginalFile | ExternalLink DataSource .NET | copy formulas between workbooks | range copy with formulas Aspose | C# workbook external reference
// Common Searches: Aspose.Cells copy range with formulas to another workbook | how to keep external links when copying cells Aspose.Cells | update external link paths after copying range .NET | PasteOptions.All with external references Aspose | change ExternalLink.DataSource programmatically
// Developer Intent: Copy a cell range that contains formulas referencing external workbooks into a new workbook and automatically retarget those external references to the new file.
// Use Cases: Create client‑specific financial reports by copying a templated calculation block while redirecting its external data source to the client’s workbook. | Generate offline analysis files by moving a data‑driven chart range from a master workbook to a standalone file and fixing the linked data paths. | Automate per‑project workbook generation where a range that pulls from a central data file is duplicated and its source link is switched to the project‑specific data source.
// AI Prompts: Write C# code that uses Aspose.Cells to copy a range containing formulas with external links from one workbook to another and then updates each ExternalLink.DataSource to a new filename. | Explain how PasteOptions.IgnoreLinksToOriginalFile affects copying formulas with external references in Aspose.Cells and outline the steps to modify those links after the copy.

using System;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsExamples
{
    // Loads a source workbook, creates an empty destination workbook, defines matching source and destination ranges, and copies the range with PasteOptions set to include all content while preserving external links. After the copy, iterates through the destination workbook's ExternalLinks collection and updates each link's DataSource to the new file name before saving.
    class CopyRangeWithExternalLinksDemo
    {
        static void Main()
        {
            try
            {
                // Path to the source workbook that contains formulas with external links
                string sourcePath = "source.xlsx";

                // Verify that the source file exists to avoid FileNotFoundException
                if (!File.Exists(sourcePath))
                {
                    Console.WriteLine($"Source file '{sourcePath}' not found.");
                    return;
                }

                // Load the source workbook
                Workbook sourceWb = new Workbook(sourcePath);

                // Create a new (empty) destination workbook
                Workbook destWb = new Workbook();

                // Add a worksheet to the destination workbook where the range will be copied
                Worksheet destSheet = destWb.Worksheets.Add("CopiedData");

                // Define the source range that includes formulas (adjust the address as needed)
                AsposeRange sourceRange = sourceWb.Worksheets[0].Cells.CreateRange("A1:C5");

                // Define the destination range (same size) in the destination worksheet
                AsposeRange destRange = destSheet.Cells.CreateRange("A1:C5");

                // Set paste options – copy everything (values, formulas, formats, etc.)
                PasteOptions pasteOptions = new PasteOptions
                {
                    PasteType = PasteType.All,
                    // Ensure links are not ignored so they are copied and can be updated later
                    IgnoreLinksToOriginalFile = false
                };

                // Copy the source range to the destination range using the provided options
                destRange.Copy(sourceRange, pasteOptions);

                // After copying, update external link references in the destination workbook
                // to point to the new file (e.g., "dest.xlsx").
                string newFileName = "dest.xlsx";
                foreach (ExternalLink link in destWb.Worksheets.ExternalLinks)
                {
                    // Update the DataSource (the external file path) to the new file name
                    link.DataSource = newFileName;
                }

                // Save the destination workbook
                destWb.Save(newFileName);
                Console.WriteLine($"Workbook saved as '{newFileName}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
