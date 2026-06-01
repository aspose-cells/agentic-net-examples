using System;
using System.IO;
using Aspose.Cells;

namespace ExportNamedRange
{
    class Program
    {
        static void Main()
        {
            try
            {
                const string sourcePath = "SourceWorkbook.xlsx";
                const string destPath = "ExportedNamedRange.xlsx";

                // Verify source file exists to avoid FileNotFoundException
                if (!File.Exists(sourcePath))
                {
                    Console.WriteLine($"Source file '{sourcePath}' not found.");
                    return;
                }

                // Load the source workbook
                Workbook sourceWorkbook = new Workbook(sourcePath);

                // Retrieve the named range (assumes the name is "MyRange")
                Name namedRange = sourceWorkbook.Worksheets.Names["MyRange"];
                if (namedRange == null)
                {
                    Console.WriteLine("Named range 'MyRange' not found.");
                    return;
                }

                // Remove leading '=' from RefersTo to obtain the address
                string rangeAddress = namedRange.RefersTo.StartsWith("=")
                    ? namedRange.RefersTo.Substring(1)
                    : namedRange.RefersTo;

                // Create a Range object for the source data
                Aspose.Cells.Range sourceRange = sourceWorkbook.Worksheets[0].Cells.CreateRange(rangeAddress);

                // Create a new workbook for the export
                Workbook destWorkbook = new Workbook();

                // Create a destination range with the same size as the source range
                Aspose.Cells.Range destRange = destWorkbook.Worksheets[0].Cells.CreateRange(
                    0, 0, sourceRange.RowCount, sourceRange.ColumnCount);

                // Set paste options to copy everything (values, formulas, formatting, etc.)
                PasteOptions pasteOptions = new PasteOptions
                {
                    PasteType = PasteType.All
                };

                // Copy the source range to the destination range preserving formatting
                destRange.Copy(sourceRange, pasteOptions);

                // Save the destination workbook
                destWorkbook.Save(destPath);

                Console.WriteLine("Named range exported successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}