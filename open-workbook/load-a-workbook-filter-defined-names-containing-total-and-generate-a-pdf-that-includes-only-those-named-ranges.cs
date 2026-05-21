using System;
using System.IO;
using System.Linq;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using CellsRange = Aspose.Cells.Range;

namespace AsposeCellsNamedRangePdf
{
    class Program
    {
        static void Main()
        {
            try
            {
                const string inputPath = "InputWorkbook.xlsx";
                const string outputPdfPath = "FilteredTotalRanges.pdf";

                // Verify that the source workbook exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                    return;
                }

                // Load the source workbook
                Workbook sourceWorkbook = new Workbook(inputPath);

                // Get all defined names (workbook‑ and worksheet‑scoped)
                Name[] allNames = sourceWorkbook.Worksheets.Names.Filter(NameScopeType.All, -1);

                // Filter names containing "Total" (case‑insensitive)
                var totalNames = allNames
                    .Where(n => n.Text.IndexOf("Total", StringComparison.OrdinalIgnoreCase) >= 0)
                    .ToArray();

                // Create a new workbook to hold the selected named ranges
                Workbook pdfWorkbook = new Workbook();
                pdfWorkbook.Worksheets.Clear(); // remove the default empty sheet

                // Copy each filtered named range into its own worksheet
                foreach (Name name in totalNames)
                {
                    // Retrieve the range(s) referenced by the name
                    CellsRange[] ranges = name.GetRanges();

                    foreach (CellsRange srcRange in ranges)
                    {
                        // Add a new worksheet named after the defined name
                        Worksheet destSheet = pdfWorkbook.Worksheets.Add(name.Text);

                        // Create a destination range with the same size as the source range, starting at A1
                        CellsRange destRange = destSheet.Cells.CreateRange(
                            0, 0, srcRange.RowCount, srcRange.ColumnCount);

                        // Copy values, formulas, and formatting
                        srcRange.Copy(destRange);
                    }
                }

                // Ensure the workbook is not empty
                if (pdfWorkbook.Worksheets.Count == 0)
                {
                    pdfWorkbook.Worksheets.Add("NoTotalRanges");
                }

                // Prepare PDF save options (customize if needed)
                PdfSaveOptions pdfOptions = new PdfSaveOptions();

                // Save the workbook as a PDF
                pdfWorkbook.Save(outputPdfPath, pdfOptions);
                Console.WriteLine($"PDF saved successfully to \"{outputPdfPath}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}