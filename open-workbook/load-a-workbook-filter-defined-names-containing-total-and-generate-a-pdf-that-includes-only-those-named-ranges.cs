// Title: C# – Filter Excel defined names containing “Total” and export the matched ranges to a PDF with Aspose.Cells
// Common Searches: Aspose.Cells C# filter defined names that include the word Total and generate a PDF | How to copy only named ranges with "Total" to a new workbook and save as PDF using Aspose.Cells | C# code to extract named ranges matching a keyword and convert them to PDF with Aspose.Cells | Create PDF from specific named ranges in Excel with Aspose.Cells .NET
// Tags: keyword‑based defined name filtering Aspose.Cells | named‑range extraction to PDF C# | copy selected ranges into new worksheet Aspose.Cells | PDF generation from specific Excel ranges .NET | C# Aspose.Cells range selection and PDF export

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook, selects all defined names whose text contains "Total" (case‑insensitive), copies each referenced range into a new worksheet, and saves the assembled workbook as a PDF using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            const string sourcePath = "SourceWorkbook.xlsx";
            const string outputPdf = "FilteredTotals.pdf";

            // Verify source file exists
            if (!File.Exists(sourcePath))
                throw new FileNotFoundException($"Source file not found: {sourcePath}");

            // Load the source workbook
            Workbook sourceWb = new Workbook(sourcePath);

            // Collect defined names that contain "Total" (case‑insensitive)
            List<Name> totalNames = new List<Name>();
            foreach (Name name in sourceWb.Worksheets.Names)
            {
                if (!string.IsNullOrEmpty(name.Text) &&
                    name.Text.IndexOf("Total", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    totalNames.Add(name);
                }
            }

            // Prepare result workbook
            Workbook resultWb = new Workbook();
            // Remove the default empty sheet
            if (resultWb.Worksheets.Count > 0)
                resultWb.Worksheets.RemoveAt(0);
            // Add a sheet to hold the filtered ranges
            Worksheet resultSheet = resultWb.Worksheets.Add("FilteredRanges");

            int nextRow = 0; // Row index where the next range will be placed

            foreach (Name totalName in totalNames)
            {
                // Resolve the range referenced by the defined name
                Aspose.Cells.Range srcRange = sourceWb.Worksheets.GetRangeByName(totalName.Text);
                if (srcRange == null)
                    continue; // Skip if the name does not refer to a valid range

                int rowCount = srcRange.RowCount;
                int colCount = srcRange.ColumnCount;

                // Optionally write the name above the copied data (skip for the first block)
                if (nextRow > 0)
                {
                    resultSheet.Cells[nextRow, 0].PutValue(totalName.Text);
                    nextRow++; // Move one row down to start the data
                }

                // Copy the source range to the result sheet at the current position
                Aspose.Cells.Range destRange = resultSheet.Cells.CreateRange(nextRow, 0, rowCount, colCount);
                destRange.Copy(srcRange);

                // Update the next row position (add one empty row after each block)
                nextRow += rowCount + 1;
            }

            // Save the result workbook as PDF
            resultWb.Save(outputPdf, SaveFormat.Pdf);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
