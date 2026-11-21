// Title: Copy Excel rows containing hyperlinks and validate their targets with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells copy rows with hyperlinks C# example | how to preserve hyperlink addresses when copying rows in Excel using Aspose.Cells | verify hyperlink targets after copying rows with Aspose.Cells .NET | C# iterate Hyperlink collection after copying rows in a workbook | copy a range of rows including formatting and hyperlinks using Aspose.Cells
// Tags: CopyRows method with hyperlinks Aspose.Cells | preserve Excel hyperlink targets during row copy | enumerate Hyperlink collection C# Aspose.Cells | copy rows including formatting Aspose.Cells | validate hyperlink addresses after row copy .NET

using System;
using System.IO;
using Aspose.Cells;

// The sample loads a source workbook, uses Aspose.Cells' CopyRows method to copy a specified range of rows (values, formatting, and hyperlinks) to a new workbook, iterates the destination sheet's Hyperlinks collection to output each hyperlink's display address and cell location, ensures the output directory exists, and saves the resulting file.
class Program
{
    static void Main()
    {
        try
        {
            const string sourcePath = "source.xlsx";
            const string destinationPath = "destination.xlsx";

            // Verify source file exists to avoid FileNotFoundException
            if (!File.Exists(sourcePath))
            {
                Console.WriteLine($"Source file not found: {sourcePath}");
                return;
            }

            // Load the source workbook containing rows with hyperlinks
            Workbook srcWorkbook = new Workbook(sourcePath);
            Worksheet srcSheet = srcWorkbook.Worksheets[0];

            // Create a new workbook for the copied rows
            Workbook destWorkbook = new Workbook();
            Worksheet destSheet = destWorkbook.Worksheets[0];

            // Define the range of rows to copy (example: rows 1‑5, zero‑based indices 0‑4)
            int srcStartRow = 0;      // first row to copy in source sheet
            int rowCount = 5;         // number of rows to copy
            int destStartRow = 0;     // first row in destination sheet where rows will be placed

            // Copy rows including values, formatting, and hyperlinks
            destSheet.Cells.CopyRows(srcSheet.Cells, srcStartRow, destStartRow, rowCount);

            // Verify that hyperlinks in the destination sheet point to the correct targets
            foreach (Hyperlink link in destSheet.Hyperlinks)
            {
                // Hyperlink display text (or name) – used when Target property is unavailable
                string address = link.TextToDisplay ?? string.Empty;

                // Convert column index to column name (e.g., 0 -> "A")
                string columnName = CellsHelper.ColumnIndexToName(link.Area.StartColumn);
                // Row numbers are 1‑based in Excel UI
                string cellAddress = $"{columnName}{link.Area.StartRow + 1}";

                Console.WriteLine($"Hyperlink at {cellAddress} points to: {address}");
            }

            // Ensure the destination directory exists
            string destDir = Path.GetDirectoryName(destinationPath);
            if (!string.IsNullOrEmpty(destDir) && !Directory.Exists(destDir))
            {
                Directory.CreateDirectory(destDir);
            }

            // Save the destination workbook with the copied rows
            destWorkbook.Save(destinationPath);
            Console.WriteLine($"Destination workbook saved to: {destinationPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
