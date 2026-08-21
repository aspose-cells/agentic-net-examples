// Title: Copy a cell range to CSV while preserving commas and quotes – Aspose.Cells for .NET (C#)
// Description: Demonstrates how to copy a defined range (A1:C3) from one workbook to another, configure TxtSaveOptions with a comma separator and QuoteType.Always, limit the export to the copied area, and save the result as a CSV file that retains delimiters and text qualifiers.
// Keywords: Aspose.Cells | C# | .NET | export range to CSV | preserve commas in CSV | preserve quotes in CSV | TxtSaveOptions | QuoteType.Always | range copy | CSV delimiter
// Common Searches: Aspose.Cells export specific range to CSV | keep commas and quotes when saving CSV with Aspose.Cells | C# copy cell range and save as CSV using Aspose | TxtSaveOptions delimiter and quoting options | how to export only a CellArea to CSV in .NET
// Developer Intent: Generate a CSV file from a selected cell range that maintains original commas and quotation marks.
// Use Cases: Create a CSV report that includes user‑entered text containing commas or quotes without breaking column alignment. | Provide a data extract for an external system that requires every field to be quoted for reliable parsing. | Export a subset of a workbook (e.g., a table or filtered data) as CSV while preserving delimiters and text qualifiers.
// AI Prompts: Write C# code using Aspose.Cells to copy a range and save it as a CSV with commas as separators and all fields quoted. | Show how to set TxtSaveOptions in Aspose.Cells to export only a specific CellArea and keep text qualifiers intact. | Provide an example that copies a range from one workbook to another and then exports the range to CSV, handling commas and quotes correctly.

using System;
using Aspose.Cells;

namespace AsposeCellsRangeToCsv
{
    // Demonstrates how to copy a defined range (A1:C3) from one workbook to another, configure TxtSaveOptions with a comma separator and QuoteType.Always, limit the export to the copied area, and save the result as a CSV file that retains delimiters and text qualifiers.
    class Program
    {
        static void Main()
        {
            try
            {
                // ---------- Create source workbook and fill sample data ----------
                Workbook sourceWb = new Workbook();
                Worksheet sourceSheet = sourceWb.Worksheets[0];
                Cells srcCells = sourceSheet.Cells;

                // Sample data with commas and quotes to test text qualifiers
                srcCells["A1"].PutValue("Name");
                srcCells["B1"].PutValue("Description");
                srcCells["C1"].PutValue("Value");

                srcCells["A2"].PutValue("Item 1");
                srcCells["B2"].PutValue("A, B, C");          // contains delimiter
                srcCells["C2"].PutValue("\"Quoted\" Text"); // contains quotes

                srcCells["A3"].PutValue("Item 2");
                srcCells["B3"].PutValue("Simple");
                srcCells["C3"].PutValue(12345);

                // ---------- Define the source range ----------
                Aspose.Cells.Range sourceRange = srcCells.CreateRange("A1:C3");

                // ---------- Create destination workbook ----------
                Workbook destWb = new Workbook();
                Worksheet destSheet = destWb.Worksheets[0];
                Cells destCells = destSheet.Cells;

                // Destination range must match the size of the source range
                Aspose.Cells.Range destRange = destCells.CreateRange("A1:C3");

                // ---------- Copy the range (including values, formats, etc.) ----------
                PasteOptions pasteOptions = new PasteOptions
                {
                    PasteType = PasteType.All // copy everything
                };
                destRange.Copy(sourceRange, pasteOptions);

                // ---------- Configure CSV (text) save options ----------
                TxtSaveOptions saveOptions = new TxtSaveOptions
                {
                    // Use comma as delimiter
                    Separator = ',',
                    // Always quote each field to preserve text qualifiers
                    QuoteType = TxtValueQuoteType.Always,
                    // Export only the copied range
                    ExportArea = new CellArea
                    {
                        StartRow = 0,
                        EndRow = 2,      // rows 0‑2 (A1:C3)
                        StartColumn = 0,
                        EndColumn = 2    // columns 0‑2 (A‑C)
                    },
                    // Ensure blank rows keep separators (optional, based on requirement)
                    KeepSeparatorsForBlankRow = true
                };

                // ---------- Save the destination workbook as CSV ----------
                string outputPath = "ExportedRange.csv";
                destWb.Save(outputPath, saveOptions);

                Console.WriteLine($"Range successfully exported to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
