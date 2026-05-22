using System;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsTableDuplicate
{
    class Program
    {
        static void Main()
        {
            try
            {
                const string sourcePath = "source.xlsx";
                const string outputPath = "output.xlsx";

                // Verify source file exists to avoid FileNotFoundException
                if (!File.Exists(sourcePath))
                {
                    Console.WriteLine($"Source file not found: {sourcePath}");
                    return;
                }

                // Load the source workbook
                Workbook workbook = new Workbook(sourcePath);

                // Assume the table to duplicate is on the first worksheet
                Worksheet sourceSheet = workbook.Worksheets[0];

                // Add a new worksheet for the duplicated table
                int newSheetIndex = workbook.Worksheets.Add();
                Worksheet destinationSheet = workbook.Worksheets[newSheetIndex];
                destinationSheet.Name = "DuplicatedTable";

                // Define the range that contains the original table.
                // Adjust the address to match the actual table range in your file.
                AsposeRange sourceRange = sourceSheet.Cells.CreateRange("A1:C5");

                // Define the destination range where the table will be copied.
                AsposeRange destinationRange = destinationSheet.Cells.CreateRange("A1:C5");

                // Set paste options to copy everything (values, formulas, formats, styles, etc.)
                PasteOptions pasteOptions = new PasteOptions
                {
                    PasteType = PasteType.All
                };

                // Perform the copy operation preserving formulas and styles.
                destinationRange.Copy(sourceRange, pasteOptions);

                // Save the workbook with the duplicated table.
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to {outputPath}");
            }
            catch (Exception ex)
            {
                // Log any unexpected errors
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}