// Title: Import data into a named range from another workbook using Aspose.Cells for .NET
// Description: C# example that loads source.xlsx and dest.xlsx, ensures the named range "MyRange" exists (creates A1:D5 on the first sheet if needed), clears its current content, copies values, formulas and formatting from the source range, and saves the result as dest_updated.xlsx.
// Keywords: Aspose.Cells import named range | C# copy range between workbooks | create named range programmatically | clear range before copy Aspose.Cells | overwrite cells safely .NET
// Common Searches: Aspose.Cells copy data to a named range in another workbook | how to create a named range if it does not exist using Aspose.Cells | clear existing cells before copying with Aspose.Cells | overwrite named range safely Aspose.Cells C#
// Developer Intent: Copy a source range into a destination named range, creating the range when missing and removing previous content.
// Use Cases: Refresh a template workbook by importing the latest dataset into a predefined named range. | Populate a financial model with master‑sheet data while guaranteeing the target range exists and is clean. | Automate dashboard updates by overwriting a named range with calculations from a separate workbook.
// AI Prompts: Write C# code with Aspose.Cells that copies a range from source.xlsx to a named range "MyRange" in dest.xlsx, creating the range if absent and clearing existing cells first. | Explain error handling for missing source or destination files when importing data into a named range with Aspose.Cells. | Provide a step‑by‑step tutorial for copying values, formulas, and formatting between workbooks using Aspose.Cells for .NET.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsImportIntoNamedRange
{
    // C# example that loads source.xlsx and dest.xlsx, ensures the named range "MyRange" exists (creates A1:D5 on the first sheet if needed), clears its current content, copies values, formulas and formatting from the source range, and saves the result as dest_updated.xlsx.
    class Program
    {
        static void Main()
        {
            try
            {
                const string sourcePath = "source.xlsx";
                const string destPath = "dest.xlsx";

                // Verify that the required files exist
                if (!File.Exists(sourcePath))
                    throw new FileNotFoundException($"Source file not found: {sourcePath}");
                if (!File.Exists(destPath))
                    throw new FileNotFoundException($"Destination file not found: {destPath}");

                // Load workbooks
                Workbook sourceWb = new Workbook(sourcePath);
                Workbook destWb = new Workbook(destPath);

                const string rangeName = "MyRange";
                Name namedRange = destWb.Worksheets.Names[rangeName];

                // Create the named range if it does not exist
                if (namedRange == null)
                {
                    // Define the address (e.g., A1:D5) on the first worksheet
                    string address = $"={destWb.Worksheets[0].Name}!$A$1:$D$5";
                    int idx = destWb.Worksheets.Names.Add(rangeName);
                    destWb.Worksheets.Names[idx].RefersTo = address;
                    namedRange = destWb.Worksheets.Names[rangeName];
                }

                // Get the Range object that the name refers to
                Aspose.Cells.Range destRange = namedRange.GetRange();

                // Clear existing contents in the destination range
                destRange.Worksheet.Cells.ClearRange(
                    destRange.FirstRow,
                    destRange.FirstColumn,
                    destRange.RowCount,
                    destRange.ColumnCount);

                // Create a source range of the same size
                Worksheet srcSheet = sourceWb.Worksheets[0];
                Aspose.Cells.Range srcRange = srcSheet.Cells.CreateRange(
                    destRange.FirstRow,
                    destRange.FirstColumn,
                    destRange.RowCount,
                    destRange.ColumnCount);

                // Copy data (values, formulas, formatting) from source to destination
                destRange.CopyData(srcRange);

                // Save the updated destination workbook
                destWb.Save("dest_updated.xlsx");
                Console.WriteLine("Data imported successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
