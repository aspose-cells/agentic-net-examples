// Title: Copy a range with formulas and external links to a new workbook using Aspose.Cells for .NET
// Description: Loads a source workbook, creates an empty destination workbook, copies a specified range (including formulas, values, and formatting) with PasteOptions that preserve external links, updates each ExternalLink.DataSource to point to the original file, and saves the result as a new Excel file.
// Keywords: Aspose.Cells copy range | copy formulas with external links | PasteOptions IgnoreLinksToOriginalFile | ExternalLink DataSource | C# Excel workbook copy | .NET Excel range transfer | update external references Aspose
// Common Searches: Aspose.Cells copy range preserving external links | C# copy cells with formulas to another workbook | how to update ExternalLink.DataSource after copying cells | PasteOptions to retain external references in Aspose.Cells | copy Excel range with formulas and links using .NET
// Developer Intent: Transfer a cell range that contains formulas with external references to a new workbook and re‑point those links to the original source file.
// Use Cases: Create a report workbook that reuses calculation blocks from a master file while keeping formulas functional. | Generate a template by copying a data‑driven area and ensuring all external links resolve to the source workbook. | Automate a summary sheet that pulls formulas from another workbook and updates link paths for consistent data sourcing.
// AI Prompts: Write C# code with Aspose.Cells to copy a range that includes formulas and external links to a new workbook, then adjust the link paths to a given file. | Explain the role of PasteOptions.IgnoreLinksToOriginalFile when copying formulas that reference external workbooks. | Provide a step‑by‑step tutorial for updating ExternalLink.DataSource after copying a range in Aspose.Cells for .NET.

using System;
using System.IO;
using Aspose.Cells;

// Loads a source workbook, creates an empty destination workbook, copies a specified range (including formulas, values, and formatting) with PasteOptions that preserve external links, updates each ExternalLink.DataSource to point to the original file, and saves the result as a new Excel file.
class Program
{
    static void Main()
    {
        try
        {
            const string sourcePath = "Source.xlsx";
            const string destinationPath = "Destination.xlsx";

            // Verify that the source file exists to avoid FileNotFoundException.
            if (!File.Exists(sourcePath))
            {
                Console.WriteLine($"Source file not found: {sourcePath}");
                return;
            }

            // Load the source workbook that contains formulas with external references.
            Workbook sourceWorkbook = new Workbook(sourcePath);

            // Create a new (empty) destination workbook.
            Workbook destinationWorkbook = new Workbook();
            // Remove the default sheet and add a fresh one for the copied data.
            destinationWorkbook.Worksheets.Clear();
            Worksheet destSheet = destinationWorkbook.Worksheets.Add("CopiedData");

            // Define the range in the source workbook that you want to copy.
            // Use fully qualified Aspose.Cells.Range to avoid ambiguity with System.Range.
            Aspose.Cells.Range sourceRange = sourceWorkbook.Worksheets[0].Cells.CreateRange("A1:B10");

            // Define the destination range with the same size.
            Aspose.Cells.Range destRange = destSheet.Cells.CreateRange("A1:B10");

            // Configure paste options to copy everything (values, formulas, formats, etc.).
            // Setting IgnoreLinksToOriginalFile to false ensures that external links are retained.
            PasteOptions pasteOptions = new PasteOptions
            {
                PasteType = PasteType.All,
                IgnoreLinksToOriginalFile = false
            };

            // Perform the copy operation.
            destRange.Copy(sourceRange, pasteOptions);

            // Update external links in the destination workbook to point to the source workbook.
            foreach (ExternalLink link in destinationWorkbook.Worksheets.ExternalLinks)
            {
                // Assign the source workbook file name (or full path) to the DataSource property.
                link.DataSource = sourceWorkbook.FileName;
            }

            // Save the destination workbook.
            destinationWorkbook.Save(destinationPath);
            Console.WriteLine($"Workbook copied successfully to {destinationPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
